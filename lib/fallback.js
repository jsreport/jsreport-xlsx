const uuid = require('uuid')
const Promise = require('bluebird')
const fs = require('fs')
const path = require('path')
const excelbuilder = require('msexcel-builder-extended')
const responseXlsx = require('./responseXlsx.js')

module.exports = (previousError, request, response) => {
  const generationId = uuid.v1()

  const result = response.content.toString()
  const workbook = excelbuilder.createWorkbook(request.reporter.options.tempDirectory, generationId + '.xlsx')

  let start = result.indexOf('<worksheet')
  let sheetNumber = 1
  while (start >= 0) {
    const worksheet = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" + result.substring(
      start,
      result.indexOf('</worksheet>', start) + '</worksheet>'.length, start)

    const startSheetName = worksheet.indexOf('name=')
    let sheetName = 'Sheet ' + sheetNumber.toString()
    if (startSheetName > 0 && startSheetName < worksheet.indexOf('>', worksheet.indexOf('<worksheet'))) {
      const s = startSheetName + 'name="'.length
      const e = worksheet.indexOf('"', s)
      sheetName = worksheet.substring(s, e)
    }
    const sheet1 = workbook.createSheet(sheetName, 0, 0)
    sheet1.raw(worksheet)
    sheetNumber = sheetNumber + 1
    start = result.indexOf('<worksheet', start + 1)
  }

  if (result.indexOf('<styleSheet') > 0) {
    const stylesheet = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" + result.substring(
      result.indexOf('<styleSheet'),
      result.indexOf('</styleSheet>') + '</styleSheet>'.length)
    workbook.st.raw(stylesheet)
  }

  return Promise.promisify(workbook.save).call(workbook).then(() => {
    response.stream = fs.createReadStream(path.join(request.reporter.options.tempDirectory, generationId + '.xlsx'))
    return responseXlsx(request, response)
  }).catch(function (e) {
    const error = new Error('Unable to parse xlsx template JSON string (maybe you are missing {{{xlsxPrint}}} at the end?): \n' + response.content.toString().substring(0, 100) + '...')
    error.weak = true
    throw error
  })
}
