import React, { Component } from 'react'
import Studio from 'jsreport-studio'

let _xlsxTemplateUploadButton

export default class ImageUploadButton extends Component {
  static propTypes = {
    tab: React.PropTypes.object,
    onUpdate: React.PropTypes.func.isRequired
  }

  // we need to have global action in main_dev which is triggered when users clicks on + on images
  // this triggers invisible button in the toolbar
  static OpenUpload (forNew = true) {
    _xlsxTemplateUploadButton.openFileDialog(forNew)
  }

  componentDidMount () {
    _xlsxTemplateUploadButton = this
  }

  upload (e) {
    if (!e.target.files.length) {
      return
    }

    const file = e.target.files[0]
    const reader = new FileReader()

    reader.onloadend = async () => {
      this.refs.file.value = ''
      if (this.forNew) {
        if (Studio.workspaces) {
          await Studio.workspaces.save()
        }

        let response = await Studio.api.post('/odata/xlsxTemplates', {
          data: {
            contentRaw: reader.result.substring(reader.result.indexOf('base64,') + 'base64,'.length),
            name: file.name.replace(/.xlsx$/, '')
          }
        })
        response.__entitySet = 'xlsxTemplates'

        Studio.addExistingEntity(response)
        Studio.openTab(Object.assign({}, response))
      } else {
        if (Studio.workspaces) {
          Studio.updateEntity({
            _id: this.props.tab.entity._id,
            contentRaw: reader.result.substring(reader.result.indexOf('base64,') + 'base64,'.length)
          })

          await Studio.workspaces.save()
        } else {
          await Studio.api.patch(`/odata/xlsxTemplates(${this.props.tab.entity._id})`, {
            data: {
              contentRaw: reader.result.substring(reader.result.indexOf('base64,') + 'base64,'.length)
            }
          })
          Studio.loadEntity(this.props.tab.entity._id, true)
        }
      }
    }

    reader.onerror = function () {
      alert('There was an error reading the file!')
    }

    reader.readAsDataURL(file)
  }

  openFileDialog (forNew) {
    this.forNew = forNew

    this.refs.file.dispatchEvent(new MouseEvent('click', {
      'view': window,
      'bubbles': false,
      'cancelable': true
    }))
  }

  renderUpload () {
    return <input
      type='file' key='file' ref='file' style={{display: 'none'}} onChange={(e) => this.upload(e)}
      accept='.xlsx'></input>
  }

  render () {
    return this.renderUpload(true)
  }
}

