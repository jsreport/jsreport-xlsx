/*eslint no-unused-vars: 1*/
/*eslint no-new-func: 0*/
/*global __rootDirectory*/
/*global m*/

(function (global) {
  function print () {
    return JSON.stringify(this.ctx.root.$xlsxTemplate)
  }

  function fixSheetDataEmptyString () {
    if (!this.ctx.data.$xlsxTemplate) {
      return
    }
    var sheet = this.ctx.data.$xlsxTemplate['xl/worksheets/sheet1.xml']
    if (!sheet) {
      return
    }

    if (!sheet.worksheet.sheetData[0].row) {
      sheet.worksheet.sheetData[0] = {
        row: []
      }
    }
  }

  function replace (filePath, path) {
    if (typeof path === 'string') {
      var holder = new Function('obj', 'return obj.' + path.split('.').slice(0, -1).join('.'))(this.ctx.root.$xlsxTemplate[filePath])
      var pathFragmentToBeReplaced = path.split('.')[path.split('.').length - 1]
      this.$replacedValue = new Function('obj', 'return obj.' + pathFragmentToBeReplaced)(holder)
      var json = xml2json(this.tagCtx.render(this.ctx.data))
      new Function('obj', 'json', 'return obj.' + pathFragmentToBeReplaced + ' = json')(holder, json)
    } else {
      this.ctx.root.$xlsxTemplate[filePath] = xml2json(this.tagCtx.render(this.ctx.data))
    }

    return ''
  }

  function remove (filePath, path, index) {
    var obj = this.ctx.root.$xlsxTemplate[filePath]
    var collection = new Function('obj', 'return obj.' + path)(obj)
    this.ctx.root.$removedItem = collection[index]
    collection.splice(index, 1)
    return ''
  }

  function merge (filePath, path) {
    var json = xml2jsonUnwrap(this.tagCtx.render(this.ctx.data))

    var mergeTarget = new Function('obj', 'return obj.' + path)(this.ctx.root.$xlsxTemplate[filePath])

    _.merge(mergeTarget, json)
    return ''
  }

  function add (filePath, xmlPath) {
    var obj = this.ctx.root.$xlsxTemplate[filePath]
    var collection = new Function('obj', 'return obj.' + xmlPath)(obj)
    collection.push(xml2jsonUnwrap(this.tagCtx.render(this.ctx.data)))
    return ''
  }

  function addSheet (name) {
    var id = this.ctx.root.$xlsxTemplate['xl/workbook.xml'].workbook.sheets.length + 1
    var fileName = 'sheet' + id
    var fileFullName = fileName + '.xml'
    var path = 'xl/worksheets/' + fileFullName

    this.ctx.root.$xlsxTemplate['[Content_Types].xml'].Types.Override.push({
      $: {
        PartName: '/' + path,
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
      }
    })
    this.ctx.root.$xlsxTemplate['xl/workbook.xml'].workbook.sheets[0].sheet.push({
      $: {
        name: name,
        sheetId: id + '',
        'r:id': fileName
      }
    })
    this.ctx.root.$xlsxTemplate['xl/_rels/workbook.xml.rels'].Relationships.Relationship.push({
      $: {
        Id: fileName,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
        Target: 'worksheets/' + fileFullName
      }
    })
    this.ctx.root.$xlsxTemplate[path] = { worksheet: xml2jsonUnwrap(this.tagCtx.render(this.ctx.data)) }
  }

  function ensureDrawingOnSheet (sheetFullName) {
    var drawingFullName
    if (this.ctx.root.$xlsxTemplate['xl/worksheets/' + sheetFullName].worksheet.drawing) {
      var rid = this.ctx.root.$xlsxTemplate['xl/worksheets/' + sheetFullName].worksheet.drawing.$['r:id']
      this.ctx.root.$xlsxTemplate['xl/worksheets/_rels/' + sheetFullName + '.rels'].Relationships.Relationship.forEach(function (r) {
        if (r.$.Id === rid) {
          drawingFullName = r.$.Target.replace('../drawings/', '')
        }
      })
    } else {
      var numberOfDrawings = 0
      this.ctx.root.$xlsxTemplate['[Content_Types].xml'].Types.Override.forEach(function (o) {
        numberOfDrawings += o.$.PartName.indexOf('/xl/drawings') === -1 ? 0 : 1
      })

      var drawingName = 'drawing' + (numberOfDrawings + 1)
      drawingFullName = drawingName + '.xml'

      this.ctx.root.$xlsxTemplate['xl/worksheets/' + sheetFullName].worksheet.drawing = {
        $: {
          'r:id': drawingName
        }
      }

      this.ctx.root.$xlsxTemplate['xl/worksheets/_rels/' + sheetFullName + '.rels'].Relationships.Relationship.push({
        $: {
          Id: drawingName,
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          Target: '../drawings/' + drawingFullName
        }
      })

      this.ctx.root.$xlsxTemplate['[Content_Types].xml'].Types.Override.push({
        $: {
          PartName: '/xl/drawings/' + drawingFullName,
          ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml'
        }
      })

      this.ctx.root.$xlsxTemplate['xl/drawings/' + drawingFullName] = {
        'xdr:wsDr': xml2jsonUnwrap(
          '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>')
      }
    }

    return drawingFullName
  }

  function ensureRelOnSheet (sheetFullName) {
    this.ctx.root.$xlsxTemplate['xl/worksheets/_rels/' + sheetFullName + '.rels'] =
      this.ctx.root.$xlsxTemplate['xl/worksheets/_rels/' + sheetFullName + '.rels'] || {
        Relationships: {
          $: {
            'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
          },
          Relationship: []
        }
      }
  }

  function addImage (imageName, sheetFullName, fromCol, fromRow, toCol, toRow) {
    var name = imageName + '.png'

    if (!this.ctx.root.$xlsxTemplate['xl/media/' + name]) {
      this.ctx.root.$xlsxTemplate['xl/media/' + name] = this.tagCtx.render(this.ctx.data)
    }

    if (!this.ctx.root.$xlsxTemplate['[Content_Types].xml'].Types.Default.filter((t) => t.$.Extension === 'png').length) {
      this.ctx.root.$xlsxTemplate['[Content_Types].xml'].Types.Default.push({
        $: {
          Extension: 'png',
          ContentType: 'image/png'
        }
      })
    }

    ensureRelOnSheet.call(this, sheetFullName)
    var drawingFullName = ensureDrawingOnSheet.call(this, sheetFullName)

    const drawingRelPath = 'xl/drawings/_rels/' + drawingFullName + '.rels'
    this.ctx.root.$xlsxTemplate[drawingRelPath] =
      this.ctx.root.$xlsxTemplate[drawingRelPath] || {
        Relationships: {
          $: { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' },
          Relationship: []
        }
      }

    const relNumber = this.ctx.root.$xlsxTemplate[drawingRelPath].Relationships.Relationship.length + 1
    const relName = 'rId' + relNumber

    if (!this.ctx.root.$xlsxTemplate[drawingRelPath].Relationships.Relationship.filter((r) => r.$.Id === imageName).length) {
      console.log('adding ' + imageName)
      this.ctx.root.$xlsxTemplate[drawingRelPath].Relationships.Relationship.push({
        $: {
          Id: relName,
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          Target: '../media/' + name
        }
      })
    }

    var drawing = this.ctx.root.$xlsxTemplate['xl/drawings/' + drawingFullName]
    drawing['xdr:wsDr']['xdr:twoCellAnchor'] = drawing['xdr:wsDr']['xdr:twoCellAnchor'] || []

    drawing['xdr:wsDr']['xdr:twoCellAnchor'].push(xml2jsonUnwrap(
      '<xdr:twoCellAnchor><xdr:from><xdr:col>' + fromCol +
      '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + fromRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>' +
      toCol + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + toRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr>' +
      '<xdr:cNvPr id="' + relNumber + '" name="Picture"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill>' +
      '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="' + relName + '"><a:extLst>' +
      '<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" ' +
      'val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" ' +
      'cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>'
    ))
  }

  function safeRequire (moduleName) {
    try {
      return require(moduleName)
    } catch (e) {
      return null
    }
  }

  var _ = safeRequire('lodash') || safeRequire(m.data.$xlsxModuleDirname + '/node_modules/lodash')
  var xml2js = safeRequire('xml2js') || safeRequire(m.data.$xlsxModuleDirname + '/node_modules/xml2js')

  var xml2jsonUnwrap = function (xml) {
    var result = xml2json(xml)
    return result[Object.keys(result)[0]]
  }

  var xml2json = function (xml) {
    var result = {}
    var err = null
    xml2js.parseString(xml, function (aerr, res) {
      result = res
      err = aerr
    })

    if (err) {
      throw err
    }

    return result
  }

  function jsrenderHandlebarsCompatibility (fn) {
    return function () {
      if (arguments.length && arguments[arguments.length - 1].name && arguments[arguments.length - 1].hash) {
        // handlebars
        var options = arguments[arguments.length - 1]

        this.ctx = {
          root: options.data.root,
          data: this
        }

        if (options.fn) {
          this.tagCtx = {
            render: options.fn
          }
        }
      } else {
        if (this.tagCtx) {
          this.ctx.data = this.tagCtx.view.data
        }
      }

      fixSheetDataEmptyString.call(this)

      return fn.apply(this, arguments)
    }
  }

  global.xlsxReplace = jsrenderHandlebarsCompatibility(replace)
  global.xlsxMerge = jsrenderHandlebarsCompatibility(merge)
  global.xlsxAdd = jsrenderHandlebarsCompatibility(add)
  global.xlsxRemove = jsrenderHandlebarsCompatibility(remove)
  global.xlsxAddImage = jsrenderHandlebarsCompatibility(addImage)
  global.xlsxAddSheet = jsrenderHandlebarsCompatibility(addSheet)
  global.xlsxPrint = jsrenderHandlebarsCompatibility(print)
})(this)
