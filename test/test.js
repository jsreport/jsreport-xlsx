import Reporter from 'jsreport-core'
import path from 'path'
import xlsx from 'xlsx'
import 'should'
import fs from 'fs'

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
})