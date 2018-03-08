const recipe = require('./recipe.js')
const serialize = require('./serialize.js')
const shortid = require('shortid')
const fs = require('fs')
const path = require('path')
const Promise = require('bluebird')
const vm = require('vm')
const responseXlsx = require('./responseXlsx.js')

const FS = Promise.promisifyAll(fs)
let defaultXlsxTemplate

function parseBoolean (v) {
  if (v && typeof v === 'string') {
    return v === 'true'
  }

  return v
}

function parseNumber (v) {
  if (v && typeof v === 'string') {
    return parseInt(v)
  }

  return v
}

module.exports = (reporter, definition) => {
  // make sure options passed through env variables are parsed
  definition.options.previewInExcelOnline = parseBoolean(definition.options.previewInExcelOnline)
  definition.options.escapeAmp = parseBoolean(definition.options.escapeAmp)
  definition.options.addBufferSize = parseNumber(definition.options.addBufferSize)
  definition.options.numberOfParsedAddIterations = parseNumber(definition.options.numberOfParsedAddIterations)

  reporter.extensionsManager.recipes.push({
    name: 'xlsx',
    execute: (req, res) => recipe(reporter, req, res)
  })

  if (reporter.compilation) {
    reporter.compilation.resource('defaultXlsxTemplate.json', path.join(__dirname, '../static', 'defaultXlsxTemplate.json'))
    reporter.compilation.resource('helpers.js', path.join(__dirname, '../static', 'helpers.js'))
  }

  reporter.options.tasks.modules.push({
    alias: 'fsproxy.js',
    path: path.join(__dirname, '../lib/fsproxy.js')
  })

  reporter.options.tasks.modules.push({
    alias: 'lodash',
    path: require.resolve('lodash')
  })

  reporter.options.tasks.modules.push({
    alias: 'xml2js',
    path: require.resolve('xml2js')
  })

  if (reporter.options.tasks.allowedModules !== '*') {
    reporter.options.tasks.allowedModules.push('path')
  }

  reporter.documentStore.registerEntityType('XlsxTemplateType', {
    _id: { type: 'Edm.String', key: true },
    'shortid': { type: 'Edm.String' },
    'name': { type: 'Edm.String', publicKey: 'true' },
    'contentRaw': { type: 'Edm.Binary', document: { extension: 'xlsx' } },
    'content': { type: 'Edm.String', document: { extension: 'txt' } }
  })

  reporter.documentStore.registerEntitySet('xlsxTemplates', {
    entityType: 'jsreport.XlsxTemplateType',
    humanReadableKey: 'shortid',
    splitIntoDirectories: true
  })

  reporter.documentStore.registerComplexType('XlsxTemplateRefType', {
    'shortid': { type: 'Edm.String' }
  })

  if (reporter.documentStore.model.entityTypes['TemplateType']) {
    reporter.documentStore.model.entityTypes['TemplateType'].xlsxTemplate = { type: 'Collection(jsreport.XlsxTemplateRefType)' }
  }

  reporter.initializeListeners.add('xlsxTemplates', () => {
    reporter.documentStore.collection('xlsxTemplates').beforeInsertListeners.add('xlsxTemplates', (doc) => {
      doc.shortid = doc.shortid || shortid.generate()
      return serialize(doc.contentRaw, reporter.options.tempAutoCleanupDirectory).then((serialized) => (doc.content = serialized))
    })

    reporter.documentStore.collection('xlsxTemplates').beforeUpdateListeners.add('xlsxTemplates', (query, update, req) => {
      if (update.$set && update.$set.contentRaw) {
        return serialize(update.$set.contentRaw, reporter.options.tempAutoCleanupDirectory).then((serialized) => (update.$set.content = serialized))
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
        return serialize(req.template.xlsxTemplate.content, reporter.options.tempAutoCleanupDirectory).then((serialized) => JSON.parse(serialized))
      }

      const docs = await reporter.documentStore.collection('xlsxTemplates').find({ shortid: req.template.xlsxTemplate.shortid }, req)
      if (!docs.length) {
        throw new Error('Unable to find xlsx template with shortid ' + req.template.xlsxTemplate.shortid)
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
