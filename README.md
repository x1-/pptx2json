# pptx2json

![Node.js CI](https://github.com/x1-/pptx2json/workflows/Node.js%20CI/badge.svg?branch=master)
![npm](https://img.shields.io/npm/v/npm)
[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

Operating Powerpoint file (Microsoft Office 2007 and later) as Office Open XML without external tools, just pure Javascript.  
Providing two main functions:
- Parse from a PowerPoint file to Json
- Parse from a Json to PowerPoint

The images, movies, audio files and so on in a PowerPoint are treated as binary.  
This is strongly inspired from [pptx-compose](https://github.com/shobhitsharma/pptx-compose).  

## Installation

```sh
$ npm install pptx2json
```

## Usage

### Parse a PowerPoint file to Json

```javascript
const PPTX2Json = require('pptx2json');
const pptx2json = new PPTX2Json();

const json = await pptx2json.toJson('path/to/pptx');
```

### Rebuild a PowerPoint from Json

If you want to get a buffer below:

```javascript
const PPTX2Json = require('pptx2json');
const pptx2json = new PPTX2Json();

const json = await pptx2json.toJson('path/to/pptx');
:
// return buffer to pptx 
const pptx = await pptx2json.toPPTX(json);
```

Otherwise want to write a file below:

```javascript
const PPTX2Json = require('pptx2json');
const pptx2json = new PPTX2Json();

const json = await pptx2json.toJson('path/to/pptx');
:
// write pptx to the path 
await pptx2json.toPPTX(json, {'file': 'path/to/output.pptx'});
```

### Get max id, rid in slides.

```javascript
const PPTX2Json = require('pptx2json');
const pptx2json = new PPTX2Json();

const json = await pptx2json.toJson(testPPTX);
const ids = pptx2json.getMaxSlideIds(json);
// {'id': 5, 'rid': 3}
```

### Get slideLayoutType Hash.

```javascript
const PPTX2Json = require('pptx2json');
const pptx2json = new PPTX2Json();

const json = await pptx2json.toJson(testPPTX);
const table = pptx2json.getSlideLayoutTypeHash(json);
// {
//    'title': 'ppt/slideLayouts/slideLayout1.xml',
//    'blank': 'ppt/slideLayouts/slideLayout7.xml'
// }
```

## Dependencies

- [jszip](https://github.com/Stuk/jszip)
- [xml2js](https://github.com/Leonidas-from-XIV/node-xml2js)


## Reference

- [PresentationML Presentation](http://officeopenxml.com/prPresentation.php)
- [PresentationML ドキュメントを操作する](https://docs.microsoft.com/ja-jp/office/open-xml/working-with-presentationml-documents)
