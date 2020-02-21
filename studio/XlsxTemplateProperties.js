import React, { Component } from 'react'
import Studio from 'jsreport-studio'

const EntityRefSelect = Studio.EntityRefSelect

export default class XlsxTemplateProperties extends Component {
  static selectItems (entities) {
    return Object.keys(entities).filter((k) => entities[k].__entitySet === 'xlsxTemplates').map((k) => entities[k])
  }

  static selectAssets (entities) {
    return Object.keys(entities).filter((k) => entities[k].__entitySet === 'assets').map((k) => entities[k])
  }

  static title (entity, entities) {
    if (
      !entity.xlsxTemplate ||
      (!entity.xlsxTemplate.shortid && !entity.xlsxTemplate.templateAssetShortid)
    ) {
      return 'xlsx template'
    }

    const foundItems = XlsxTemplateProperties.selectItems(entities).filter((e) => entity.xlsxTemplate.shortid === e.shortid)
    const foundAssets = XlsxTemplateProperties.selectAssets(entities).filter((e) => entity.xlsxTemplate.templateAssetShortid === e.shortid)

    if (!foundItems.length && !foundAssets.length) {
      return 'xlsx template'
    }

    let name

    if (foundAssets.length) {
      name = foundAssets[0].name
    } else {
      name = foundItems[0].name
    }

    return 'xlsx template: ' + name
  }

  componentDidMount () {
    this.removeInvalidXlsxTemplateReferences()
  }

  componentDidUpdate () {
    this.removeInvalidXlsxTemplateReferences()
  }

  removeInvalidXlsxTemplateReferences () {
    const { entity, entities, onChange } = this.props

    if (!entity.xlsxTemplate) {
      return
    }

    const updatedXlsxTemplates = Object.keys(entities).filter((k) => entities[k].__entitySet === 'xlsxTemplates' && entities[k].shortid === entity.xlsxTemplate.shortid)
    const updatedXlsxAssets = Object.keys(entities).filter((k) => entities[k].__entitySet === 'assets' && entities[k].shortid === entity.xlsxTemplate.templateAssetShortid)

    const newXlsxTemplate = { ...entity.xlsxTemplate }
    let changed = false

    if (entity.xlsxTemplate.shortid && updatedXlsxTemplates.length === 0) {
      changed = true
      newXlsxTemplate.shortid = null
    }

    if (entity.xlsxTemplate.templateAssetShortid && updatedXlsxAssets.length === 0) {
      changed = true
      newXlsxTemplate.templateAssetShortid = null
    }

    if (changed) {
      onChange({ _id: entity._id, xlsxTemplate: Object.keys(newXlsxTemplate).length ? newXlsxTemplate : null })
    }
  }

  changeXlsxTemplate (oldXlsxTemplate, prop, value) {
    let newValue

    if (value == null) {
      newValue = { ...oldXlsxTemplate }
      newValue[prop] = null
    } else {
      return { ...oldXlsxTemplate, [prop]: value }
    }

    newValue = Object.keys(newValue).length ? newValue : null

    return newValue
  }

  render () {
    const { entity, onChange } = this.props

    return (
      <div className='properties-section'>
        <div className='form-group'>
          <label>xlsx asset</label>
          <EntityRefSelect
            headingLabel='Select docx template'
            value={entity.xlsxTemplate ? entity.xlsxTemplate.templateAssetShortid : ''}
            onChange={(selected) => onChange({
              _id: entity._id,
              xlsxTemplate: this.changeXlsxTemplate(entity.xlsxTemplate, 'templateAssetShortid', selected != null && selected.length > 0 ? selected[0].shortid : null)
            })}
            filter={(references) => ({ data: references.assets })}
          />
        </div>
        <div className='form-group'>
          <label>xlsx template (deprecated)</label>
          <EntityRefSelect
            headingLabel='Select xlsx template'
            filter={(references) => ({ xlsxTemplates: references.xlsxTemplates })}
            value={entity.xlsxTemplate ? entity.xlsxTemplate.shortid : null}
            onChange={(selected) => onChange({
              _id: entity._id,
              xlsxTemplate: this.changeXlsxTemplate(entity.xlsxTemplate, 'shortid', selected != null && selected.length > 0 ? selected[0].shortid : null)
            })}
          />
        </div>
      </div>
    )
  }
}
