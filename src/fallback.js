import uuid from 'uuid'
import Promise from 'bluebird'
import fs from 'fs'
import path from 'path'
import excelbuilder from 'msexcel-builder-extended'
import responseXlsx from './responseXlsx.js'

export default function fallback (previousError, request, response) {
  var generationId = uuid.v1()

  var result = response.content.toString()
  var workbook = excelbuilder.createWorkbook(request.reporter.options.tempDirectory, generationId + '.xlsx')

  var start = result.indexOf('<worksheet')
  var sheetNumber = 1
  while (start >= 0) {
    var worksheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + result.substring(
        start,
        result.indexOf('</worksheet>', start) + '</worksheet>'.length, start)

    var startSheetName = worksheet.indexOf('name=')
    var sheetName = 'Sheet ' + sheetNumber.toString()
    if (startSheetName > 0 && startSheetName < worksheet.indexOf('>', worksheet.indexOf('<worksheet'))) {
      var s = startSheetName + 'name="'.length
      var e = worksheet.indexOf('"', s)
      sheetName = worksheet.substring(s, e)
    }
    var sheet1 = workbook.createSheet(sheetName, 0, 0)
    sheet1.raw(worksheet)
    sheetNumber = sheetNumber + 1
    start = result.indexOf('<worksheet', start + 1)
  }

  if (result.indexOf('<styleSheet') > 0) {
    var stylesheet = '<?xml version="1.0" encoding="UTF-8"" standalone="yes"?>' + result.substring(
        result.indexOf('<styleSheet'),
        result.indexOf('</styleSheet>') + '</styleSheet>'.length)
    workbook.st.raw(stylesheet)
  }

  return Promise.promisify(workbook.save).call(workbook).then(() => {
    response.stream = fs.createReadStream(path.join(request.reporter.options.tempDirectory, generationId + '.xlsx'))
    return responseXlsx(request, response)
  }).catch(function (e) {
    var error = new Error('Unable to parse xlsx template JSON string (maybe you are missing {{{xlsxPrint}}} at the end?): \n' + response.content.toString().substring(0, 100) + '...')
    error.weak = true
    throw error
  })
};
