const fs = require('fs')
const Promise = require('bluebird')
const archiver = require('archiver')
const fallback = require('./fallback.js')
const responseXlsx = require('./responseXlsx.js')
const jsonToXml = require('./jsonToXml.js')
const stream = require('stream')
const zlib = require('zlib')
const merge = require('merge2')

const stringToStream = (str) => {
  var s = new stream.Readable()
  s.push(str)
  s.push(null)
  return s
}

module.exports = async (reporter, req, res) => {
  reporter.logger.debug('Parsing xlsx content', req)

  const contentString = res.content.toString()

  let $xlsxTemplate
  let $files

  try {
    let content = JSON.parse(contentString)
    $xlsxTemplate = content.$xlsxTemplate
    $files = content.$files
  } catch (e) {
    reporter.logger.warn('Unable to parse xlsx template JSON string (maybe you are missing {{{xlsxPrint}}} at the end?): \n' + contentString.substring(0, 100) + '... \n' + e.stack, req)
    return fallback(e, reporter, req, res)
  }

  const files = Object.keys($xlsxTemplate).map((k) => {
    if (k.includes('xl/media/') || k.includes('.bin')) {
      return {
        path: k,
        data: Buffer.from($xlsxTemplate[k], 'base64')
      }
    }

    if (k.includes('.xml')) {
      const xmlAndFiles = jsonToXml($xlsxTemplate[k])
      const fullXml = xmlAndFiles.xml

      if (fullXml.indexOf('&&') < 0) {
        return {
          path: k,
          data: Buffer.from(fullXml, 'utf8')
        }
      }

      const xmlStream = merge()

      if (fullXml.indexOf('&&') < 0) {
        return {
          path: k,
          data: Buffer.from(fullXml, 'utf8')
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
      data: Buffer.from($xlsxTemplate[k], 'utf8')
    }
  })

  const {
    pathToFile: xlsxFileName,
    stream: output
  } = await reporter.writeTempFileStream((uuid) => `${uuid}.xlsx`)

  return new Promise((resolve, reject) => {
    reporter.logger.debug(`Zipping prepared xml files into ${xlsxFileName}`, req)

    const archive = archiver('zip')

    output.on('close', () => {
      reporter.logger.debug('Successfully zipped now.', req)
      res.stream = fs.createReadStream(xlsxFileName)
      responseXlsx(reporter, req, res).then(() => resolve()).catch((e) => reject(e))
    })

    archive.on('error', (err) => reject(err))

    archive.pipe(output)

    files.forEach((f) => archive.append(f.data, { name: f.path }))

    archive.finalize()
  })
}
