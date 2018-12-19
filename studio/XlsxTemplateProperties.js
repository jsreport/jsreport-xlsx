import React, { Component } from 'react'
import Studio from 'jsreport-studio'

const EntityRefSelect = Studio.EntityRefSelect

export default class XlsxTemplateProperties extends Component {
  static selectItems (entities) {
    return Object.keys(entities).filter((k) => entities[k].__entitySet === 'xlsxTemplates').map((k) => entities[k])
  }

  static title (entity, entities) {
    if (!entity.xlsxTemplate || !entity.xlsxTemplate.shortid) {
      return 'xlsx template'
    }

    const foundItems = XlsxTemplateProperties.selectItems(entities).filter((e) => entity.xlsxTemplate.shortid === e.shortid)

    if (!foundItems.length) {
      return 'xlsx template'
    }

    return 'xlsx template: ' + foundItems[0].name
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

    if (updatedXlsxTemplates.length === 0) {
      onChange({ _id: entity._id, xlsxTemplate: null })
    }
  }

  render () {
    const { entity, onChange } = this.props

    return (
      <div className='properties-section'>
        <div className='form-group'>
          <EntityRefSelect
            headingLabel='Select xlsx template'
            filter={(references) => ({ xlsxTemplates: references.xlsxTemplates })}
            value={entity.xlsxTemplate ? entity.xlsxTemplate.shortid : null}
            onChange={(selected) => onChange({ _id: entity._id, xlsxTemplate: selected != null && selected.length > 0 ? { shortid: selected[0].shortid } : null })}
          />
        </div>
      </div>
    )
  }
}
