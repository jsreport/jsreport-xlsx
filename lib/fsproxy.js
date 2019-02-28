const uuid = require('uuid')
const path = require('path')
const fs = require('fs')
const mkdirp = require('mkdirp')
const zlib = require('zlib')

module.exports.write = (tmp, data) => {
  const file = path.join(tmp, uuid.v1() + '.xml')

  mkdirp.sync(tmp)

  fs.writeFileSync(file, zlib.deflateSync(data))
  return file
}
