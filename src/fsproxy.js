import uuid from 'uuid'
import path from 'path'
import fs from 'fs'
import zlib from 'zlib'

module.exports.write = (tmp, data) => {
  const file = path.join(tmp, uuid.v1() + '.xml')

  fs.writeFileSync(file, zlib.deflateSync(data))
  return file
}
