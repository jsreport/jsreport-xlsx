# jsreport-xlsx
[![NPM Version](http://img.shields.io/npm/v/jsreport-xlsx.svg?style=flat-square)](https://npmjs.com/package/jsreport-xlsx)

jsreport recipe rendering excels directly from open xml

See http://jsreport.net/learn/xlsx

##Installation

> **npm install jsreport-xlsx**
##Usage
To use `recipe` in for template rendering set `template.recipe=xlsx` in the rendering request.

```js
{
  template: { content: '...', recipe: 'xlsx', enginne: '...' }
}
```

##jsreport-core
You can apply this extension also manually to [jsreport-core](https://github.com/jsreport/jsreport-core)

```js
var jsreport = require('jsreport-core')()
jsreport.use(require('jsreport-xlsx')())
```
