import fs from 'fs'
import path from 'path'
import Promise from 'bluebird'
import archiver from 'archiver'
import uuid from 'uuid'
import fallback from './fallback.js'
import responseXlsx from './responseXlsx.js'
import jsonToXml from './jsonToXml.js'
import xml2js from 'xml2js'

export default function (req, res) {
  const contentString = res.content.toString()
  let content
  try {
    content = JSON.parse(contentString)
  } catch (e) {
    req.logger.warn('Unable to parse xlsx template JSON string (maybe you are missing {{{xlsxPrint}}} at the end?): \n' + contentString.substring(0, 100) + '... \n' + e.stack)
    return fallback(e, req, res)
  }

  const files = []
  Object.keys(content).forEach((k) => {
    var buf = null
    if (k.includes('.xml')) {
      buf = new Buffer(jsonToXml(content[k]), 'utf8')
    }

    if (k.includes('xl/media/') || k.includes('.bin')) {
      buf = new Buffer(content[k], 'base64')
    }

    if (!buf) {
      buf = new Buffer(content[k], 'utf8')
    }

    files.push({
      path: k,
      data: buf
    })
  })

  return new Promise((resolve, reject) => {
    const id = uuid.v1()
    const xlsxFileName = path.join(req.reporter.options.tempDirectory, id + '.xlsx')
    const archive = archiver('zip')
    const output = fs.createWriteStream(xlsxFileName)

    output.on('close', function () {
      req.logger.debug('Successfully zipped now.')
      res.stream = fs.createReadStream(xlsxFileName)
      responseXlsx(req, res).then(() => resolve()).catch((e) => reject(e))
    })

    archive.on('error', (err) => reject(err))

    archive.pipe(output)
    files.forEach((f) => archive.append(f.data, { name: f.path }))
    archive.finalize()
  })
}