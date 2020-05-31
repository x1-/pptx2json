/**
 * @module pptx2json
 * @fileoverview Convert Open Office XML pptx buffer to JSON and ecoding XML
 *
 * @author x1- <viva008@gmail.com>
 */

'use strict';

const fs = require('fs').promises;
const os = require('os');
const path = require('path');
const assert = require('assert');
const JSZip = require('jszip');
const xml2js = require('xml2js');


/**
 * @class PPTX2Json
 *
 */
class PPTX2Json {
  /**
   * @method constructor
   * 
   * constructor for PPTX2Json
   * 
   * @param {Object} options {
   *   'jszipBinary': nodebuffer(default) | blob | arraybuffer | uint8array | nodestream, see: https://stuk.github.io/jszip/documentation/api_jszip/support.html
   *   'jszipGenerateType': nodebuffer(default) | blob | arraybuffer | uint8array | nodestream
   * }
   * @returns PPTX2Json
   */
  constructor(options) {
    this.options = options || {};
  }

  async jszip2json(jszip) {
    const json = {};
    await Promise.all(
      Object.keys(jszip.files).map(async relativePath => {
        const file = jszip.file(relativePath);
        const ext = path.extname(relativePath);

        let content;
        if (!file || file.dir) {
          return;
        } else if (ext === '.xml' || ext === '.rels') {
          const xml = await file.async("string");
          content = await xml2js.parseStringPromise(xml);
        } else {
          content = await file.async(this.options['jszipBinary'] || 'nodebuffer');  // images, audio files, movies, etc.
        }
        json[relativePath] = content;  
      })
    );
    return json;
  }

  /**
   * @method toJson
   * 
   * Parse PowerPoint file to Json.
   * 
   * @param {string} file Give a path of PowerPoint.
   * @returns {Promise} json 
   */
  async toJson(file) {
    assert.equal(typeof file, 'string', "argument 'file' must be a string");

    const content = {};

    const buff = await fs.readFile(file);
    const zip = await JSZip().loadAsync(buff);

    return await this.jszip2json(zip);
  }

  json2jszip(json) {
    const zip = new JSZip();
    Object.keys(json).forEach(relativePath => {
      const ext = path.extname(relativePath);
      if (ext === '.xml' || ext === '.rels') {
        const builder = new xml2js.Builder({
          renderOpts: {
            pretty: false
          }
        });
        const xml = (builder.buildObject(json[relativePath]));
        zip.file(relativePath, xml);
      } else {
        zip.file(relativePath, json[relativePath]);
      }
    });
    return zip;
  }

  /**
   * @method toPPTX
   * 
   * Convert json to pptx.
   * It is available to add or delete slides, media before call this method.
   * 
   * @param {Object} json created from PowerPoint XMLs.
   * @param {Object} options {
   *   'file': If you want to write file, please give the output path. If not, return buffer.
   * }
   * @returns {Promise} Buffer if not file, otherwise empty.
   */
  async toPPTX(json, options) {
    options = options || {};

    assert.equal(typeof json, 'object', "argument 'json' must be an object");
    assert.equal(typeof options, 'object', "argument 'options' must be an object");

    const zip = this.json2jszip(json);

    let buf = await zip.generateAsync({
      type: this.options['jszipGenerateType'] || 'nodebuffer'
    });
    if (!options.file) {
      return buf;
    }
    return await fs.writeFile(options.file, buf);
  }
}

module.exports = PPTX2Json;
