import Promise from 'bluebird'
import httpRequest from 'request'
import toArray from 'stream-to-array'
const toArrayAsync = Promise.promisify(toArray)

const preview = (request, response) => {
  return new Promise((resolve, reject) => {
    const req = httpRequest.post('http://jsreport.net/temp', (err, resp, body) => {
      if (err) {
        return reject(err)
      }
      response.content = new Buffer('<iframe style="height:100%;width:100%" src="https://view.officeapps.live.com/op/view.aspx?src=' +
        encodeURIComponent('http://jsreport.net/temp/' + body) + '" />')
      response.headers['Content-Type'] = 'text/html'
      // sometimes files is not completely flushed and excel online cannot find it immediately
      setTimeout(function () {
        resolve()
      }, 500)
    })

    var form = req.form()
    form.append('file', response.stream)
    response.headers['Content-Type'] = 'text/html'
  })
}

export default function responseXlsx (request, response) {
  if (request.options.preview) {
    return preview(request, response)
  }

  response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetf'
  response.headers['Content-Disposition'] = 'inline; filename="report.xlsx"'
  response.headers['File-Extension'] = 'xlsx'

  return toArrayAsync(response.stream).then((buf) => {
    response.content = Buffer.concat(buf)
  })
}