import XlsxEditor from './XlsxEditor.js'
import XlsxUploadButton from './XlsxUploadButton.js'
import Properties from './XlsxTemplateProperties.js'
import Studio from 'jsreport-studio'

Studio.addEntitySet({ name: 'xlsxTemplates', faIcon: 'fa-file-excel-o', visibleName: 'xlsx template', onNew: XlsxUploadButton.OpenUpload })
Studio.addEditorComponent('xlsxTemplates', XlsxEditor)
Studio.addToolbarComponent(XlsxUploadButton)
Studio.addPropertiesComponent(Properties.title, Properties, (entity) => entity.__entitySet === 'templates' && entity.recipe === 'xlsx')
