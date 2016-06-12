import React, { Component } from 'react'
import XlsxUploadButton from './XlsxUploadButton.js'
import fileSaver from 'filesaver.js-npm'

const b64toBlob = (b64Data, contentType = '', sliceSize = 512) => {
  const byteCharacters = atob(b64Data)
  const byteArrays = []

  for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
    const slice = byteCharacters.slice(offset, offset + sliceSize)

    const byteNumbers = new Array(slice.length)
    for (let i = 0; i < slice.length; i++) {
      byteNumbers[i] = slice.charCodeAt(i)
    }

    const byteArray = new Uint8Array(byteNumbers)

    byteArrays.push(byteArray)
  }

  const blob = new Blob(byteArrays, { type: contentType })
  return blob
}

export default class ImageEditor extends Component {
  static propTypes = {
    entity: React.PropTypes.object.isRequired
  }

  download () {
    const blob = b64toBlob(this.props.entity.contentRaw, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    fileSaver.saveAs(blob, this.props.entity.name)
  }

  render () {
    const { entity } = this.props

    return (<div className='custom-editor'>
      <div><h1><i className='fa fa-file-excel-o' /> {entity.name}</h1></div>
      <div>
        <button className='button confirmation' onClick={() => this.download()}>
          <i className='fa fa-download' /> Download xlsx template
        </button>
        <button className='button confirmation' onClick={() => XlsxUploadButton.OpenUpload(false)}>
          <i className='fa fa-upload' /> Upload (edit) xlsx template
        </button>
      </div>
    </div>)
  }
}

