'use strict'

const fs = require("fs");
const JSZip = require('jszip');
const PPTX2Json = require('./index');

const testPPTX = "./fixtures/test.pptx";
const testZip = "./fixtures/test.zip";
const testImage = "./fixtures/cube.jpeg";

test('When given presentation.xml has slId, getMaxSlideIds returns max id and max rid.', () => {
  const pptx2json = new PPTX2Json();
  const json = {
    'ppt/presentation.xml': {"p:presentation": {"p:sldIdLst": [
      {"p:sldId": [{
        "$": {
          "id": "262",
          "r:id": "rId5"
        }
      },{
        "$": {
          "id": "261",
          "r:id": "rId6"
        }
      },{
        "$": {
          "id": "267",
          "r:id": "rId7"
        }
      }]}
    ]}}
  };
  const ids = pptx2json.getMaxSlideIds(json);
  expect(ids.id).toBe(267);
  expect(ids.rid).toBe(7);
});

test('When give presentation.xml does not have sldId, getMaxSlideIds returns -1.', () => {
  const pptx2json = new PPTX2Json();
  const json = {
    'ppt/presentation.xml': {"p:presentation": {"p:sldIdLst": []}}
  };
  const ids = pptx2json.getMaxSlideIds(json);
  expect(ids.id).toBe(-1);
  expect(ids.rid).toBe(-1);
});

test('When give valid pptx object, getSlideLayoutTypeHash returns valid LayoutTypeHash.', async () => {
  const pptx2json = new PPTX2Json();
  const json = await pptx2json.toJson(testPPTX);
  const table = pptx2json.getSlideLayoutTypeHash(json);
  expect(table.title).toBe('ppt/slideLayouts/slideLayout1.xml');
  expect(table.blank).toBe('ppt/slideLayouts/slideLayout7.xml');
});

test('When give valid zip object, jszip2json returns valid json.', async () => {
  const buff = fs.readFileSync(testZip);
  const zip = await JSZip().loadAsync(buff);

  const pptx2json = new PPTX2Json();
  const json = await pptx2json.jszip2json(zip);

  expect(Object.keys(json).length).toBe(3);
});

test('When give valid pptx object, toJson returns valid json.', async () => {
  const pptx2json = new PPTX2Json();
  const json = await pptx2json.toJson(testPPTX);

  expect('ppt/presentation.xml' in json).toBe(true);
});

test('When give valid json, json2jszip returns valid zip.', () => {
  const pptx2json = new PPTX2Json();
  const json = {
    'apple.xml': {
      "fruits": {
        "fruit": [
          {
            "name": "apple",
            "color": "red"
          }
        ]
      }
    }
  };
  const jszip = pptx2json.json2jszip(json);
  const files = Object.keys(jszip.files);

  expect(files.length).toBe(1);
  expect(jszip.file('apple.xml').dir).toBe(false);
});

test('When give valid pptx object, call toJson and then call toPPTX return valid pptx.', async () => {
  const pptx2json = new PPTX2Json();
  const json = await pptx2json.toJson(testPPTX);
  const pptx = await pptx2json.toPPTX(json);

  expect(pptx).toEqual(expect.anything());
});

test('When give valid pptx object, call toJson and add jpeg, then call toPPTX return valid pptx.', async () => {
  const pptx2json = new PPTX2Json();
  const json = await pptx2json.toJson(testPPTX);

  const image = fs.readFileSync(testImage);
  json['ppt/media/image6.jpeg'] = image;

  const pptx = await pptx2json.toPPTX(json);

  expect(pptx).toEqual(expect.anything());
});
