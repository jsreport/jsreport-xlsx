const recipe = require('./recipe.js')
const serialize = require('./serialize.js')
const fs = require('fs')
const path = require('path')
const Promise = require('bluebird')
const vm = require('vm')
const extend = require('node.extend.without.arrays')
const { response } = require('jsreport-office')

const FS = Promise.promisifyAll(fs)
let defaultXlsxTemplate

module.exports = (reporter, definition) => {
  if (reporter.options.xlsx) {
    reporter.logger.warn('xlsx root configuration property is deprecated. Use office property instead')
  }

  definition.options = extend(true, { preview: {} }, reporter.options.xlsx, reporter.options.office, definition.options)
  reporter.options.xlsx = definition.options

  if (definition.options.previewInExcelOnline != null) {
    reporter.logger.warn('extensions.xlsx.previewInExcelOnline configuration property is deprecated. Use office.preview.enabled=false instead')
    definition.options.preview.enabled = definition.options.previewInExcelOnline
  }

  if (definition.options.showExcelOnlineWarning != null) {
    reporter.logger.warn('extensions.xlsx.showExcelOnlineWarning configuration property is deprecated. Use office.preview.showWarning=false instead')
    definition.options.preview.showWarning = definition.options.showExcelOnlineWarning
  }

  if (definition.options.publicUriForPreview != null) {
    reporter.logger.warn('extensions.xlsx.publicUriForPreview configuration property is deprecated. Use office.preview.publicUri=https://... instead')
    definition.options.preview.publicUri = definition.options.publicUriForPreview
  }

  reporter.extensionsManager.recipes.push({
    name: 'xlsx',
    execute: (req, res) => recipe(reporter, definition, req, res)
  })

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
    reporter.documentStore.model.entityTypes['TemplateType'].xlsxTemplate = { type: 'jsreport.XlsxTemplateRefType', schema: { type: 'null' } }
  }

  reporter.initializeListeners.add('xlsxTemplates', () => {
    if (reporter.express) {
      reporter.express.exposeOptionsToApi(definition.name, {
        preview: {
          enabled: definition.options.preview.enabled,
          showWarning: definition.options.preview.showWarning
        }
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

        return FS.readFileAsync(path.join(__dirname, '../static/defaultXlsxTemplate.json')).then((content) => JSON.parse(content))
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

    const helpersScript = await FS.readFileAsync(path.join(__dirname, '../static/helpers.js'), 'utf8')

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

  reporter.on('express-configure', (app) => {
    app.get('/xlsxTemplates/office/:id/content', async (req, res) => {
      try {
        const xlsxTemplate = await reporter.documentStore.collection('xlsxTemplates').findOne({
          _id: req.params.id
        })

        if (!xlsxTemplate) {
          return res.status(404).end(`xlsxTemplate with _id "${req.params.id}" does not exists`)
        }

        req.options = req.options || {}
        req.options.preview = true

        res.meta = res.meta || {}

        await response({
          previewOptions: definition.options.preview,
          officeDocumentType: 'xlsx',
          buffer: xlsxTemplate.contentRaw
        }, req, res)

        res.setHeader('Content-Type', res.meta.contentType)

        res.end(res.content)
      } catch (e) {
        reporter.logger.warn(`Unable to get xlsxTemplate content ${e.stack}`)
        res.status(500).end(e.message)
      }
    })
  })
}
