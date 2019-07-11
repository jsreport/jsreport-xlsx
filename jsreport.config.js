const office = require('jsreport-office')

const schema = {
  type: 'object',
  properties: {
    previewInExcelOnline: { type: 'boolean' },
    publicUriForPreview: { type: 'string' },
    escapeAmp: { type: 'boolean' },
    addBufferSize: { type: 'number' },
    numberOfParsedAddIterations: { type: 'number' },
    showExcelOnlineWarning: { type: 'boolean', default: true }
  }
}

module.exports = {
  'name': 'xlsx',
  'main': 'lib/index.js',
  'optionsSchema': office.extendSchema('xlsx', {
    xlsx: { ...schema },
    extensions: {
      xlsx: { ...schema }
    }
  }),
  'dependencies': ['templates', 'data']
}
