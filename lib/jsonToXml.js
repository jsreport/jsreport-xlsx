const convertAttributes = (obj) => {
  var xml = ''
  if (obj.$) {
    for (var attrKey in obj.$) {
      xml += ` ${attrKey}="${convertEntities(obj.$[attrKey])}"`
    }
  }

  return xml
}

const entityMap = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#x27;',
  '/': '&#x2F;',
  '=': '&#x3D;'
}

const convertEntities = (str) => {
  if (!str) {
    return str
  }

  return str.replace(/[<>&'"](?!(amp;|lt;|gt;|quot;|#x27;|#x2F;|#x3D;))/g, (s) => {
    return entityMap[s]
  })
}

module.exports = (o) => {
  let files = []

  const convertBody = (obj) => {
    if (obj == null) {
      return ''
    }

    if (typeof obj === 'string') {
      return convertEntities(obj.toString())
    }

    let xml = ''
    for (let key in obj) {
      if (obj[key] == null || key === '$') {
        continue
      }

      let body = ''

      if (Array.isArray(obj[key])) {
        for (let i = 0; i < obj[key].length; i++) {
          if (obj[key][i].$$ != null) {
            files.push(obj[key][i].$$)
            xml += '&&'
            continue
          }

          if (Object.keys(obj[key][i]).length > 1 || Object.keys(obj[key][i])[0] !== '$') {
            body = convertBody(obj[key][i])
            xml += '<' + key + convertAttributes(obj[key][i])
            xml += body ? (`>${body}</${key}>\n`) : '/>'
          } else {
            xml += `<${key}${convertAttributes(obj[key][i])}/>`
          }
        }

        continue
      }

      if (key === '_') {
        return obj[key].toString()
      }

      xml += '<' + key

      xml += convertAttributes(obj[key])
      body = convertBody(obj[key])
      xml += body ? (`>${body}</${key}>`) : '/>'
    }

    return xml
  }

  return {
    xml: '<?xml version="1.0" encoding="UTF-8"?>\n' + convertBody(o),
    files: files
  }
}
