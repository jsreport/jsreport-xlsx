
const schema = {
  type: 'object',
  properties: {
    previewInExcelOnline: { type: 'boolean', $exposeToApi: true },
    publicUriForPreview: { type: 'string' },
    escapeAmp: { type: 'boolean' },
    addBufferSize: { type: 'number' },
    numberOfParsedAddIterations: { type: 'number' },
    showExcelOnlineWarning: { type: 'boolean', default: true, $exposeToApi: true }
  }
}

module.exports = {
  'name': 'xlsx',
  'main': 'lib/index.js',
  'optionsSchema': {
    xlsx: { ...schema },
    extensions: {
      xlsx: { ...schema }
    }
  },
  'dependencies': ['templates', 'data']
}
