# jsreport-xlsx

[![NPM Version](http://img.shields.io/npm/v/jsreport-xlsx.svg?style=flat-square)](https://npmjs.com/package/jsreport-xlsx)
[![Build Status](https://travis-ci.org/jsreport/jsreport-xlsx.png?branch=master)](https://travis-ci.org/jsreport/jsreport-xlsx)

> jsreport recipe which renders excel reports based on uploaded excel templates by modifying the xlsx source using predefined templating engine helpers

See the docs http://jsreport.net/learn/xlsx

<iframe src='https://playground.jsreport.net/studio/workspace/rJftqRaQ/10?embed=1' width="100%" height="400" frameborder="0"></iframe>

##Installation

>npm install jsreport-excel



##jsreport-core

```js
var jsreport = require('jsreport-core')()
jsreport.use(require('jsreport-excel')())

```