import Reporter from 'jsreport-core'
import path from 'path'
import xlsx from 'xlsx'
import 'should'
import fs from 'fs'
import _ from 'lodash'

describe('excel recipe', () => {
  let reporter

  beforeEach((done) => {
    reporter = new Reporter({
      rootDirectory: path.join(__dirname, '../')
    })

    reporter.init().then(function () {
      done()
    }).fail(done)
  })

  const test = (contentName, assertion) => {
    return (done) => {
      reporter.render({
        template: {
          recipe: 'xlsx',
          engine: 'handlebars',
          content: fs.readFileSync(path.join(__dirname, 'content', contentName), 'utf8')
        }
      }).then((res) => {
        assertion(xlsx.read(res.content))
        done()
      }).catch(done)
    }
  }

  it('should generate empty excel by default', test('empty.handlebars', (workbook) => {
    workbook.SheetNames.should.have.length(1)
    workbook.SheetNames[0].should.be.eql('Sheet1')
  }))

  it('xlsxMerge rename-sheet', test('rename-sheet.handlebars', (workbook) => {
    workbook.SheetNames.should.have.length(1)
    workbook.SheetNames[0].should.be.eql('XXX')
  }))

  it('xlsxAdd add-row', test('add-row.handlebars', (workbook) => {
    workbook.Sheets.Sheet1.A1.should.be.ok
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
      xlsx.read(res.content).Sheets.Sheet1.A1.should.be.ok
      xlsx.read(res.content).Sheets.Sheet1.A1000.should.be.ok
    })
  })

  it('xlsxReplace replace-sheet', test('replace-sheet.handlebars', (workbook) => {
    workbook.Sheets.Sheet1.A1.should.be.ok
  }))

  it('add-sheet', test('add-sheet.handlebars', (workbook) => {
    workbook.Sheets.Test.A1.should.be.ok
  }))

  it('should be able to use uploaded xlsx template', (done) => {
    let templateContent = fs.readFileSync(path.join(__dirname, 'Book1.xlsx')).toString('base64')
    reporter.documentStore.collection('xlsxTemplates').insert({
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
        done()
      })
    }).catch(done)
  })

  it('should return iframe in preview', (done) => {
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
      done()
    }).catch(done)
  })

  it('should disable preview if the options has previewInExcelOnline === false', (done) => {
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
      done()
    }).catch(done)
  })

  it('should be able to use string helpers', (done) => {
    reporter.render({
      template: {
        recipe: 'xlsx',
        engine: 'handlebars',
        helpers: 'function foo() { return "<c><v>11</v></c>" }',
        content: fs.readFileSync(path.join(__dirname, 'content', 'add-row-block-helper.handlebars'), 'utf8')
      }
    }).then((res) => {
      xlsx.read(res.content).Sheets.Sheet1.A1.v.should.be.eql(11)
      done()
    }).catch(done)
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
      xlsx.read(res.content).Sheets.Sheet1.A1.should.be.ok
      xlsx.read(res.content).Sheets.Sheet1.A1000.should.be.ok
    })
  })
})

describe('excel recipe with in process helpers', () => {
  let reporter

  beforeEach((done) => {
    reporter = new Reporter({
      rootDirectory: path.join(__dirname, '../'),
      tasks: { strategy: 'in-process' }
    })

    reporter.init().then(function () {
      done()
    }).fail(done)
  })

  it('should be able to use native helpers', (done) => {
    reporter.render({
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
      done()
    }).catch(done)
  })
})
