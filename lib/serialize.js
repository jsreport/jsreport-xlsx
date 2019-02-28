const decompress = require('./decompress')
const Promise = require('bluebird')
const xml2js = require('xml2js-preserve-spaces')

const parseString = Promise.promisify(xml2js.parseString)

module.exports = async (base64content) => {
  const buf = Buffer.from(base64content, 'base64')
  let result = {}
  const files = await decompress()(buf)

  await Promise.all(files.map((f) => {
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

  return JSON.stringify(result)
}
