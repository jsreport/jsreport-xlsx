const recipe = require('./recipe.js')
const serialize = require('./serialize.js')
const fs = require('fs')
const path = require('path')
const Promise = require('bluebird')
const vm = require('vm')

const FS = Promise.promisifyAll(fs)
let defaultXlsxTemplate

module.exports = (reporter, definition) => {
  definition.options = Object.assign({}, reporter.options.xlsx, definition.options)
  reporter.options.xlsx = definition.options

  reporter.extensionsManager.recipes.push({
    name: 'xlsx',
    execute: (req, res) => recipe(reporter, req, res)
  })

  if (reporter.compilation) {
    reporter.compilation.resource('defaultXlsxTemplate.json', path.join(__dirname, '../static', 'defaultXlsxTemplate.json'))
    reporter.compilation.resource('helpers.js', path.join(__dirname, '../static', 'helpers.js'))
  }

  reporter.options.templatingEngines.modules.push({
    alias: 'fsproxy.js',
    path: path.join(__dirname, '../lib/fsproxy.js')
  })

  reporter.options.templatingEngines.modules.push({
    alias: 'lodash',
    path: require.resolve('lodash')
  })

  reporter.options.templatingEngines.modules.push({
    alias: 'xml2js-preserve-spaces',
    path: require.resolve('xml2js-preserve-spaces')
  })

  if (reporter.options.templatingEngines.allowedModules !== '*') {
    reporter.options.templatingEngines.allowedModules.push('path')
  }

  reporter.documentStore.registerEntityType('XlsxTemplateType', {
    'name': { type: 'Edm.String', publicKey: 'true' },
    'contentRaw': { type: 'Edm.Binary', document: { extension: 'xlsx' } },
    'content': { type: 'Edm.String', document: { extension: 'txt' } }
  })

  reporter.documentStore.registerEntitySet('xlsxTemplates', {
    entityType: 'jsreport.XlsxTemplateType',
    splitIntoDirectories: true
  })

  reporter.documentStore.registerComplexType('XlsxTemplateRefType', {
    'shortid': { type: 'Edm.String' }
  })

  if (reporter.documentStore.model.entityTypes['TemplateType']) {
    reporter.documentStore.model.entityTypes['TemplateType'].xlsxTemplate = { type: 'jsreport.XlsxTemplateRefType' }
  }

  reporter.initializeListeners.add('xlsxTemplates', () => {
    if (reporter.express) {
      reporter.express.exposeOptionsToApi(definition.name, {
        previewInExcelOnline: definition.options.previewInExcelOnline,
        showExcelOnlineWarning: definition.options.showExcelOnlineWarning
      })
    }

    reporter.documentStore.collection('xlsxTemplates').beforeInsertListeners.add('xlsxTemplates', (doc) => {
      return serialize(doc.contentRaw).then((serialized) => (doc.content = serialized))
    })

    reporter.documentStore.collection('xlsxTemplates').beforeUpdateListeners.add('xlsxTemplates', (query, update, req) => {
      if (update.$set && update.$set.contentRaw) {
        return serialize(update.$set.contentRaw).then((serialized) => (update.$set.content = serialized))
      }
    })
  })

  reporter.beforeRenderListeners.insert({ after: 'data' }, 'xlsxTemplates', async (req) => {
    if (req.template.recipe !== 'xlsx') {
      return
    }

    const findTemplate = async () => {
      if (!req.template.xlsxTemplate || (!req.template.xlsxTemplate.shortid && !req.template.xlsxTemplate.content)) {
        if (defaultXlsxTemplate) {
          return Promise.resolve(defaultXlsxTemplate)
        }

        if (reporter.execution) {
          return Promise.resolve(JSON.parse(reporter.execution.resource('defaultXlsxTemplate.json').toString()))
        }

        return FS.readFileAsync(path.join(__dirname, '../static', 'defaultXlsxTemplate.json')).then((content) => JSON.parse(content))
      }

      if (req.template.xlsxTemplate.content) {
        return serialize(req.template.xlsxTemplate.content).then((serialized) => JSON.parse(serialized))
      }

      const docs = await reporter.documentStore.collection('xlsxTemplates').find({ shortid: req.template.xlsxTemplate.shortid }, req)

      if (!docs.length) {
        throw reporter.createError(`Unable to find xlsx template with shortid ${req.template.xlsxTemplate.shortid}`, {
          statusCode: 404
        })
      }

      return JSON.parse(docs[0].content)
    }

    const template = await findTemplate()

    req.data = req.data || {}
    req.data.$xlsxTemplate = template
    req.data.$xlsxModuleDirname = path.join(__dirname, '../')
    req.data.$tempAutoCleanupDirectory = reporter.options.tempAutoCleanupDirectory
    req.data.$addBufferSize = definition.options.addBufferSize || 50000000
    req.data.$escapeAmp = definition.options.escapeAmp
    req.data.$numberOfParsedAddIterations = definition.options.numberOfParsedAddIterations == null ? 50 : definition.options.numberOfParsedAddIterations

    let helpersScript
    if (reporter.execution) {
      helpersScript = reporter.execution.resource('helpers.js')
    } else {
      helpersScript = await FS.readFileAsync(path.join(__dirname, '../', 'static', 'helpers.js'), 'utf8')
    }

    if (req.template.helpers && typeof req.template.helpers === 'object') {
      // this is the case when the jsreport is used with in-process strategy
      // and additinal helpers are passed as object
      // in this case we need to merge in xlsx helpers
      req.template.helpers.require = require
      req.template.helpers.fsproxy = require(path.join(__dirname, 'fsproxy.js'))
      return vm.runInNewContext(helpersScript, req.template.helpers)
    }

    req.template.helpers = helpersScript + '\n' + (req.template.helpers || '')
  })
}
