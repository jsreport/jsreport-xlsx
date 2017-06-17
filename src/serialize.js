import decompress from './decompress'
import Promise from 'bluebird'
import xml2js from 'xml2js'

const parseString = Promise.promisify(xml2js.parseString)

export default function (base64content, tempDirectory) {
  const buf = new Buffer(base64content, 'base64')
  let result = {}
  return decompress()(buf).then((files) => {
    return Promise.all(files.map((f) => {
      if (f.path.includes('.xml')) {
        return parseString(f.data.toString()).then((r) => (result[f.path] = r))
      }

      if (f.path.includes('xl/media/') || f.path.includes('.bin')) {
        result[f.path] = f.data.toString('base64')
      }

      if (!result[f.path]) {
        result[f.path] = f.data.toString('utf8')
      }
    }))
  }).then(() => (JSON.stringify(result)))
}
