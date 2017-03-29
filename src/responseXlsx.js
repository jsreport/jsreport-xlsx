import Promise from 'bluebird'
import httpRequest from 'request'
import toArray from 'stream-to-array'
const toArrayAsync = Promise.promisify(toArray)

const preview = (request, response, options) => {
  return new Promise((resolve, reject) => {
    const req = httpRequest.post(options.publicUriForPreview || 'http://jsreport.net/temp', (err, resp, body) => {
      if (err) {
        return reject(err)
      }
      var iframe = '<iframe style="height:100%;width:100%" src="https://view.officeapps.live.com/op/view.aspx?src=' +
        encodeURIComponent((options.publicUriForPreview || 'http://jsreport.net/temp') + '/' + body) + '" />'
      var title = request.template.name || 'jsreport'
      var html = '<html><head><title>' + title + '</title><body>' + iframe + '</body></html>'
      response.content = new Buffer(html)
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
  let options = request.reporter.options.xlsx || {}
  if (request.options.preview && options.previewInExcelOnline !== false) {
    return preview(request, response, options)
  }

  response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  response.headers['Content-Disposition'] = 'inline; filename="report.xlsx"'
  response.headers['File-Extension'] = 'xlsx'

  return toArrayAsync(response.stream).then((buf) => {
    response.content = Buffer.concat(buf)
  })
}
