import React, { Component } from 'react'

export default class Properties extends Component {
  static selectItems (entities) {
    return Object.keys(entities).filter((k) => entities[k].__entitySet === 'xlsxTemplates').map((k) => entities[k])
  }

  static title (entity, entities) {
    if (!entity.xlsxTemplate || !entity.xlsxTemplate.shortid) {
      return 'xlsx template'
    }

    const foundItems = Properties.selectItems(entities).filter((e) => entity.xlsxTemplate.shortid === e.shortid)

    if (!foundItems.length) {
      return 'xlsx template'
    }

    return 'xlsx template: ' + foundItems[0].name
  }

  render () {
    const { entity, entities, onChange } = this.props
    const xlsxTemplateItems = Properties.selectItems(entities)

    return (
      <div className='properties-section'>
        <div className='form-group'>
          <select
            value={entity.xlsxTemplate ? entity.xlsxTemplate.shortid : ''}
            onChange={(v) => onChange({_id: entity._id, xlsxTemplate: v.target.value !== 'empty' ? { shortid: v.target.value } : null})}>
            <option key='empty' value='empty'>- not selected -</option>
            {xlsxTemplateItems.map((e) => <option key={e.shortid} value={e.shortid}>{e.name}</option>)}
          </select>
        </div>
      </div>
    )
  }
}

