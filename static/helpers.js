/*eslint no-unused-vars: 1*/
/*eslint no-new-func: 0*/
/*global __rootDirectory*/
/*global m*/

(function (global) {
  function print () {
    ensureWorksheetOrder(this.ctx.root.$xlsxTemplate)
    return JSON.stringify(this.ctx.root.$xlsxTemplate)
  }

  var worksheetOrder = {
    dimension: 0,
    sheetViews: 1,
    sheetFormatPr: 2,
    cols: 3,
    sheetData: 4,
    sheetCalcPr: 5,
    sheetProtection: 6,
    protectedRanges: 7,
    scenarios: 8,
    autoFilter: 9,
    sortState: 10,
    dataConsolidate: 11,
    customSheetViews: 12,
    mergeCells: 13,
    phoneticPr: 14,
    conditionalFormatting: 15,
    dataValidations: 16,
    hyperlinks: 17,
    printOptions: 18,
    pageMargins: 19,
    pageSetup: 20,
    headerFooter: 21,
    rowBreaks: 22,
    colBreaks: 23,
    customProperties: 24,
    cellWatches: 25,
    ignoredErrors: 26,
    smartTags: 27,
    drawing: 28,
    picture: 29,
    oleObjects: 30,
    controls: 31,
    webPublishItems: 32,
    tableParts: 33,
    extLst: 34
  }

  function ensureWorksheetOrder(data) {
    for (var key in data) {
      if (key.indexOf('xl/worksheets/') !== 0) {
        continue
      }

      if (!data[key] || !data[key].worksheet) {
        continue
      }

      var worksheet = data[key].worksheet
      var sortedWorksheet = {}
      Object.keys(worksheet).sort(function (a, b) {
        return worksheetOrder[a] && worksheetOrder[b] && (worksheetOrder[a] > worksheetOrder[b])
      }).forEach(function (a) {
        sortedWorksheet[a] = worksheet[a]
      })
      data[key].worksheet = sortedWorksheet
    }
  }

  function replace (filePath, path) {
    if (typeof path === 'string') {
      var holder = new Function('obj', 'return obj.' + path.split('.').slice(0, -1).join('.'))(this.ctx.root.$xlsxTemplate[filePath])
      var pathFragmentToBeReplaced = path.split('.')[path.split('.').length - 1]
      this.$replacedValue = new Function('obj', 'return obj.' + pathFragmentToBeReplaced)(holder)
      var contentToReplace = this.tagCtx.render(this.ctx.data)
      try {
        contentToReplace = xml2jsonUnwrap(contentToReplace)
      } catch(e) {
        // not xml, it is ok, put it as the string value inside
      }
      new Function('obj', 'contentToReplace', 'return obj.' + pathFragmentToBeReplaced + ' = contentToReplace')(holder, contentToReplace)
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
    var collection = safeGet(obj, xmlPath)
    collection.push(xml2jsonUnwrap(this.tagCtx.render(this.ctx.data)))
    return ''
  }

  /**
   * Safely go through object path and create the missing object parts with
   * empty array or object to be compatible with xml -> json represantation
   */
  function safeGet(obj, path) {
    var originalObj = obj
    var paths = path.replace('[', '.').replace(']', '').split('.')

    var previous = {}
    for (var i = 0; i < paths.length; i++) {
      var objReference = 'obj["' + paths[i] + '"]'
      //if the next accessor is to array, we initialize missing parths as array, otherwise as object
      var emptySafe = ((i === paths.length  - 1) || !isNaN(paths[i + 1])) ? '[]' : '{}'
      new Function('obj', objReference + ' = ' + objReference + ' || ' + emptySafe)(obj)
      obj = new Function('obj', 'return ' + objReference)(obj)
      previous = obj
    }

    return obj
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
