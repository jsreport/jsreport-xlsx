import Reporter from 'jsreport-core'
import path from 'path'
import xlsx from 'xlsx'
import should from 'should'
import fs from 'fs'
import _ from 'lodash'

process.env.DEBUG = ''

describe('excel recipe', () => {
  let reporter

  beforeEach(() => {
    reporter = new Reporter({
      rootDirectory: path.join(__dirname, '../')
    })

    return reporter.init()
  })

  const test = (contentName, assertion) => {
    return () => reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', contentName), 'utf8')
      }
    }).then((res) => {
      assertion(xlsx.read(res.content))
    })
  }

  it('should generate empty excel by default', test('empty.handlebars', (workbook) => {
    workbook.SheetNames.should.have.length(1)
    workbook.SheetNames[0].should.be.eql('Sheet1')
  }))

  it('xlsxMerge rename-sheet', test('rename-sheet.handlebars', (workbook) => {
    workbook.SheetNames.should.have.length(1)
    workbook.SheetNames[0].should.be.eql('XXX')
  }))

  it('xlsxMerge rename-sheet-complex-path', test('rename-sheet-complex-path.handlebars', (workbook) => {
    workbook.SheetNames.should.have.length(1)
    workbook.SheetNames[0].should.be.eql('XXX')
  }))

  it('xlsxAdd add-row', test('add-row.handlebars', (workbook) => {
    workbook.Sheets.Sheet1.A1.should.be.ok()
  }))

  it('xlsxAdd add-row-complex-path', test('add-row-complex-path.handlebars', (workbook) => {
    workbook.Sheets.Sheet1.A1.should.be.ok()
  }))

  it('xlsxRemove remove-row', test('remove-row.handlebars', (workbook) => {
    should(workbook.Sheets.Sheet1.A1).not.be.ok()
  }))

  it('xlsxRemove remove-row-complex-path', test('remove-row-complex-path.handlebars', (workbook) => {
    should(workbook.Sheets.Sheet1.A1).not.be.ok()
  }))

  it('xlsxReplace replace-row', test('replace-row.handlebars', (workbook) => {
    workbook.Sheets.Sheet1.A1.v.should.be.eql('xxx')
  }))

  it('xlsxReplace replace-row-complex-path', test('replace-row-complex-path.handlebars', (workbook) => {
    workbook.Sheets.Sheet1.A1.v.should.be.eql('xxx')
  }))

  it('xlsxAdd add many row', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-many-rows.handlebars'), 'utf8')
      },
      data: {
        numbers: _.range(0, 1000)
      }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.should.be.ok()
      xlsx.read(res.content).Sheets.Sheet1.A1000.should.be.ok()
    })
  })

  it('xlsxReplace replace-sheet', test('replace-sheet.handlebars', (workbook) => {
    workbook.Sheets.Sheet1.A1.should.be.ok()
  }))

  it('add-sheet', test('add-sheet.handlebars', (workbook) => {
    workbook.Sheets.Test.A1.should.be.ok()
  }))

  it('should be able to use uploaded xlsx template', () => {
    let templateContent = fs.readFileSync(path.join(__dirname, 'Book1.xlsx')).toString('base64')
    return reporter.documentStore.collection('xlsxTemplates').insert({
      contentRaw: templateContent,
      shortid: 'foo',
      name: 'foo'
    }).then(() => {
      return reporter.render({
        template: {
          recipe: 'xlsx',
          engine: 'handlebars',
          xlsxTemplate: {
            shortid: 'foo'
          },
          content: '{{{xlsxPrint}}}'
        }
      }).then((res) => {
        let workbook = xlsx.read(res.content)
        workbook.Sheets.Sheet1.A1.v.should.be.eql(1)
      })
    })
  })

  it('should return iframe in preview', () => {
    return reporter.render({
      options: {
        preview: true
      },
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: '{{{xlsxPrint}}}'
      }
    }).then((res) => {
      res.content.toString().should.containEql('iframe')
    })
  })

  it('should disable preview if the options has previewInExcelOnline === false', () => {
    reporter.options.xlsx = { previewInExcelOnline: false }
    return reporter.render({
      options: {
        preview: true
      },
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: '{{{xlsxPrint}}}'
      }
    }).then((res) => {
      res.content.toString().should.not.containEql('iframe')
    })
  })

  it('should be able to use string helpers', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        helpers: 'function foo() { return "<c><v>11</v></c>" }',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-row-block-helper.handlebars'), 'utf8')
      }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.v.should.be.eql(11)
    })
  })

  it('should escape amps by default', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-row-with-foo-value.handlebars'), 'utf8').replace('{{foo}}', '& {{foo}} &amp;amp;')
      },
      data: { foo: '&' }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.v.should.be.eql('& & &')
    })
  })

  it('should pass escaped entities', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-row-with-foo-value.handlebars'), 'utf8').replace('{{foo}}', '{{{foo}}} > " ' + '"' + ' /')
      },
      data: { foo: '&lt;' }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.v.should.be.eql('< > " ' + '"' + ' /')
    })
  })

  it('should escape entities', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-row-with-foo-value.handlebars'), 'utf8')
      },
      data: { foo: '<' }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.v.should.be.eql('<')
    })
  })

  it('should escape entities in attributes', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', 'rename-sheet.handlebars'), 'utf8')
      },
      data: { foo: '<' }
    }).then((res) => {
      xlsx.read(res.content).Sheets['&lt;XXX'].should.be.ok()
    })
  })
})

describe('excel recipe with disabled add parsing', () => {
  let reporter

  beforeEach(() => {
    reporter = new Reporter({
      rootDirectory: path.join(__dirname, '../'),
      xlsx: {
        numberOfParsedAddIterations: 0,
        addBufferSize: 200
      }
    })

    return reporter.init()
  })

  it('should be add row', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-many-rows.handlebars'), 'utf8')
      },
      data: {
        numbers: _.range(0, 1000)
      }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.should.be.ok()
      xlsx.read(res.content).Sheets.Sheet1.A1000.should.be.ok()
    })
  })

  it('should escape amps', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-row-with-foo-value.handlebars'), 'utf8').replace('{{foo}}', '& {{foo}} &amp;')
      },
      data: { foo: '&' }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.v.should.be.eql('& & &')
    })
  })

  it('should escape entities', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-row-with-foo-value.handlebars'), 'utf8').replace('{{foo}}', '& < > " ' + "'" + ' /')
      },
      data: { }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.v.should.be.eql('& < > " ' + "'" + ' /')
    })
  })
})

describe('excel recipe with in process helpers', () => {
  let reporter

  beforeEach(() => {
    reporter = new Reporter({
      rootDirectory: path.join(__dirname, '../'),
      tasks: { strategy: 'in-process' }
    })

    return reporter.init()
  })

  it('should be able to use native helpers', () => {
    return reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        helpers: {
          foo: () => '<c><v>11</v></c>'
        },
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-row-block-helper.handlebars'), 'utf8')
      }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.v.should.be.eql(11)
    })
  })
})
