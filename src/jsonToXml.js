import os from 'os'

const convertAttributes = (obj) => {
  var xml = ''
  if (obj.$) {
    for (var attrKey in obj.$) {
      xml += ` ${attrKey}="${obj.$[attrKey]}"`
    }
  }

  return xml
}

const convertBody = (obj) => {
  if (obj == null) {
    return ''
  }

  if (typeof obj === 'string') {
    return obj.toString()
  }

  var xml = ''
  for (var key in obj) {
    if (obj[key] == null || key === '$') {
      continue
    }

    if (Array.isArray(obj[key])) {
      for (var i = 0; i < obj[key].length; i++) {
        if (Object.keys(obj[key][i]).length > 1 || Object.keys(obj[key][i])[0] !== '$') {
          var body = convertBody(obj[key][i])
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
    var body = convertBody(obj[key])
    xml += body ? (`>${body}</${key}>`) : '/>'
  }

  return xml
}

module.exports = (o) => ('<?xml version="1.0" encoding="UTF-8"?>\n' + convertBody(o))