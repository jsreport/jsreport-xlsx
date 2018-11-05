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
  static OpenUpload (forNew = true, options) {
    _xlsxTemplateUploadButton.openFileDialog(forNew, options)
  }

  componentDidMount () {
    _xlsxTemplateUploadButton = this
  }

  upload (e) {
    if (!e.target.files.length) {
      return
    }

    const xlsxDefaults = e.target.xlsxDefaults

    delete e.target.xlsxDefaults

    const file = e.target.files[0]
    const reader = new FileReader()

    reader.onloadend = async () => {
      this.refs.file.value = ''
      if (this.forNew) {
        if (Studio.workspaces) {
          await Studio.workspaces.save()
        }

        let xlsx = {}

        if (xlsxDefaults != null) {
          xlsx = Object.assign(xlsx, xlsxDefaults)
        }

        xlsx = Object.assign(xlsx, {
          contentRaw: reader.result.substring(reader.result.indexOf('base64,') + 'base64,'.length),
          name: file.name.replace(/.xlsx$/, '')
        })

        let response = await Studio.api.post('/odata/xlsxTemplates', {
          data: xlsx
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

  openFileDialog (forNew, options = {}) {
    this.forNew = forNew

    if (options.defaults) {
      this.refs.file.xlsxDefaults = options.defaults
    } else {
      delete this.refs.file.xlsxDefaults
    }

    this.refs.file.dispatchEvent(new MouseEvent('click', {
      'view': window,
      'bubbles': false,
      'cancelable': true
    }))
  }

  renderUpload () {
    return <input
      type='file' key='file' ref='file' style={{display: 'none'}} onChange={(e) => this.upload(e)}
      accept='.xlsx' />
  }

  render () {
    return this.renderUpload(true)
  }
}
