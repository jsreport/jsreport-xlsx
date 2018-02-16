const uuid = require('uuid')
const path = require('path')
const fs = require('fs')
const zlib = require('zlib')

module.exports.write = (tmp, data) => {
  const file = path.join(tmp, uuid.v1() + '.xml')

  fs.writeFileSync(file, zlib.deflateSync(data))
  return file
}
