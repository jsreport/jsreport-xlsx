import fs from 'fs'
import path from 'path'
import Promise from 'bluebird'
import archiver from 'archiver'
import uuid from 'uuid'
import fallback from './fallback.js'
import responseXlsx from './responseXlsx.js'
import jsonToXml from './jsonToXml.js'
import stream from 'stream'
import zlib from 'zlib'
import merge from 'merge2'

const stringToStream = (str) => {
  var s = new stream.Readable()
  s.push(str)
  s.push(null)
  return s
}

export default function (req, res) {
  req.logger.debug('Parsing xlsx content')
  const contentString = res.content.toString()
  let $xlsxTemplate
  let $files
  try {
    let content = JSON.parse(contentString)
    $xlsxTemplate = content.$xlsxTemplate
    $files = content.$files
  } catch (e) {
    req.logger.warn('Unable to parse xlsx template JSON string (maybe you are missing {{{xlsxPrint}}} at the end?): \n' + contentString.substring(0, 100) + '... \n' + e.stack)
    return fallback(e, req, res)
  }

  const files = Object.keys($xlsxTemplate).map((k) => {
    if (k.includes('xl/media/') || k.includes('.bin')) {
      return {
        path: k,
        data: new Buffer($xlsxTemplate[k], 'base64')
      }
    }

    if (k.includes('.xml')) {
      const xmlAndFiles = jsonToXml($xlsxTemplate[k])
      const fullXml = xmlAndFiles.xml

      if (fullXml.indexOf('&&') < 0) {
        return {
          path: k,
          data: new Buffer(fullXml, 'utf8')
        }
      }

      const xmlStream = merge()

      if (fullXml.indexOf('&&') < 0) {
        return {
          path: k,
          data: new Buffer(fullXml, 'utf8')
        }
      }

      let xml = fullXml

      while (xml) {
        const separatorIndex = xml.indexOf('&&')

        if (separatorIndex < 0) {
          xmlStream.add(stringToStream(xml))
          xml = ''
          continue
        }

        xmlStream.add(stringToStream(xml.substring(0, separatorIndex)))
        xmlStream.add(fs.createReadStream($files[xmlAndFiles.files.shift()]).pipe(zlib.createInflate()))
        xml = xml.substring(separatorIndex + '&&'.length)
      }

      return {
        path: k,
        data: xmlStream
      }
    }

    return {
      path: k,
      data: new Buffer($xlsxTemplate[k], 'utf8')
    }
  })

  return new Promise((resolve, reject) => {
    const id = uuid.v1()
    const xlsxFileName = path.join(req.reporter.options.tempDirectory, id + '.xlsx')
    req.logger.debug('Zipping prepared xml files into ' + xlsxFileName)
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