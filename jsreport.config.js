
const schema = {
  type: 'object',
  properties: {
    previewInExcelOnline: { type: 'boolean' },
    publicUriForPreview: { type: 'string' },
    escapeAmp: { type: 'boolean' },
    addBufferSize: { type: 'number' },
    numberOfParsedAddIterations: { type: 'number' }
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
