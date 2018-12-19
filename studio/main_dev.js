import XlsxEditor from './XlsxEditor.js'
import XlsxUploadButton from './XlsxUploadButton.js'
import XlsxTemplateProperties from './XlsxTemplateProperties.js'
import Studio from 'jsreport-studio'

Studio.addEntitySet({
  name: 'xlsxTemplates',
  faIcon: 'fa-file-excel-o',
  visibleName: 'xlsx template',
  onNew: (options = {}) => XlsxUploadButton.OpenUpload(true, options),
  entityTreePosition: 500
})
Studio.addEditorComponent('xlsxTemplates', XlsxEditor)
Studio.addToolbarComponent(XlsxUploadButton)
Studio.addPropertiesComponent(XlsxTemplateProperties.title, XlsxTemplateProperties, (entity) => entity.__entitySet === 'templates' && entity.recipe === 'xlsx')

Studio.previewListeners.push((request, entities) => {
  if (request.template.recipe !== 'xlsx') {
    return
  }

  if (Studio.extensions.xlsx.options.showExcelOnlineWarning === false) {
    return
  }

  if (Studio.extensions.xlsx.options.previewInExcelOnline === false) {
    return
  }

  if (Studio.getSettingValueByKey('xlsx-preview-informed', false) === true) {
    return
  }

  Studio.setSetting('xlsx-preview-informed', true)

  Studio.openModal(() => <div>
    We need to upload your excel report to our publicly hosted server to be able to use
    Excel Online Service for previewing here in the studio. You can disable it in the configuration, see <a
      href='https://github.com/jsreport/jsreport-xlsx' target='_blank'>https://github.com/jsreport/jsreport-xlsx</a> for details.
  </div>)
})
