const fs = require("fs");
const path = require("path");
const Stream = require("stream");
const zlib = require("zlib");
const crypto = require("crypto");

// third-pard dependcies
const ejs = require("../lib/ejs");
const Hzip = require("../lib/hzip");
const xml2json = require("../lib/xml2json");
const xmldom = require("../lib/xmldom");
const qr = require("../lib/qr-image");


// subpackage dependencies
const ejs4xlx = require("./ejs4xlx");
const {isType,isObject,isArray,isFunction,isString}=require('./is-type');
const {charToNum,charPlus,str2Xml,replaceLast}=require('./utils');
const {Promise_fromCallback,Promise_fromStandard,Promise_sleep}=require('./async.js');
const {renderExcel}=require('./data');


const DOMParser = xmldom.DOMParser;
const existsAsync = Promise_fromCallback(fs.exists, fs);
const readFileAsync = Promise_fromStandard(fs.readFile, fs);
const inflateRawAsync = Promise_fromStandard(zlib.inflateRaw, zlib);

const render = async function(buffer, filter, _data_, hzip, options) {
  var anonymous, buffer2, data, entries, flt, l, len1, sharedStrings2, str, updateEntryAsync;
  if (hzip === void 0 || hzip === null) {
    hzip = new Hzip(buffer);
  }
  updateEntryAsync = Promise_fromStandard(hzip.updateEntry, hzip);
  entries = hzip.entries;
  data = {
    _data_: _data_
  };
  data._charPlus_ = charPlus;
  data._charToNum_ = charToNum;
  data._str2Xml_ = str2Xml;
  data._acVar_ = {
    _ss_len: 0
  };
  sharedStrings2 = [Buffer.from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">")];
  data._ps_ = function(val) {
    var index;
    val = str2Xml(val);
    sharedStrings2.push(Buffer.from("<si><t xml:space=\"preserve\">" + val + "</t></si>"));
    index = data._acVar_._ss_len;
    data._acVar_._ss_len++;
    return String(index);
  };
  for (l = 0, len1 = filter.length; l < len1; l++) {
    flt = filter[l];
    if (!flt.notEjs) {
      str = ejs.parse(flt.buffer);
      anonymous = eval("(async function anonymous(_args) {\n" + str + "\n})");
      buffer2 = (await anonymous.call(this, data));
    } else {
      buffer2 = flt.buffer;
    }
    await updateEntryAsync.apply(hzip, [flt.path, buffer2]);
  }
  sharedStrings2.push(Buffer.from("</sst>"));
  buffer2 = Buffer.concat(sharedStrings2);
  await updateEntryAsync.apply(hzip, ["xl/sharedStrings.xml", buffer2]);
  return hzip.buffer;
};

const sheetSufStr = Buffer.from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<%\nvar _data_ = _args._data_;\nvar _charPlus_ = _args._charPlus_;\nvar _charToNum_ = _args._charToNum_;\nvar _str2Xml_ = _args._str2Xml_;\nvar _hideSheet_ = _args._hideSheet_;\nvar _showSheet_ = _args._showSheet_;\nvar _deleteSheet_ = _args._deleteSheet_;\nvar _ps_ = _args._ps_;\nvar _pi_ = _args._pi_;\nvar _pf_ = _args._pf_;\nvar _acVar_ = _args._acVar_;\nvar _r = 0;\nvar _c = 0;\nvar _row = 0;\nvar _col = \"\";\nvar _rc = \"\";\nvar _img_ = _args._img_;\nvar _qrcode_ = _args._qrcode_;\nvar _mergeCellArr_ = [];\nvar _mergeCellFn_ = function(mclStr) {\n	_mergeCellArr_.push(mclStr);\n};\nvar _hyperlinkArr_ = [];\nvar _outlineLevel_ = _args._outlineLevel_;\n%>");

const sharedStrings2Prx = Buffer.from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">");

var xjOp = {
  object: true,
  reversible: true,
  coerce: false,
  trim: false,
  sanitize: false
};



/**
 * render excel with a callback style
 * @param {Buffer} exlBuf 
 * @param {*} _data_ 
 * @param {Function} callback 
 */
function renderExcelCb(exlBuf, _data_, callback) {
  return renderExcel(exlBuf, _data_)
    .then(function (buf2) {
      return callback(null, buf2);
    })
    .catch(function (err) {
      return callback(err);
    });
};


var renderPath = async function(ejsDir, data) {
  var buffer, config, configPath, exists, exlBuf, extname, filter, ftObj, key, l, len1, obj, val;
  configPath = ejsDir + "/config.json";
  exists = (await existsAsync(configPath));
  if (exists === false) {
    extname = path.extname(ejsDir);
    if (extname === "") {
      ejsDir = ejsDir + ".xlsx";
    }
    exlBuf = (await readFileAsync(ejsDir));
    return (await renderExcel(exlBuf, data));
  }
  config = (await readFileAsync(configPath, "utf8"));
  config = JSON.decode(config);
  filter = [];
  for (l = 0, len1 = config.length; l < len1; l++) {
    obj = config[l];
    key = obj.key;
    val = obj.value;
    if (key === void 0 || key === null) {
      continue;
    }
    extname = path.extname(val);
    ftObj = {
      path: val
    };
    if (extname !== ".xml" && extname !== ".rels") {
      ftObj.notEjs = true;
    }
    ftObj.buffer = (await readFileAsync(ejsDir + "/" + key));
    filter.push(ftObj);
  }
  buffer = (await readFileAsync(ejsDir + "/" + path.basename(ejsDir) + ".zip"));
  return (await render(buffer, filter, data));
};


var getExcelArr  = async function(buffer) {
  var buf, cEle, crStr, cs, enr, ens, entries, entry, fileName, hzip, i, l, len1, len2, len3, len4, m, n, numcr, numcrArr, o, p, ref2, ref3, row, sharedJson, sharedStr, sheet, sheetArr, sheetStr, sheets, sheetsEns, sir, vStr, vStr2;
  sharedStr = null;
  sheets = [];
  hzip = new Hzip(buffer);
  entries = hzip.entries;
  for (l = 0, len1 = entries.length; l < len1; l++) {
    entry = entries[l];
    fileName = entry.fileName;
    if (fileName === "xl/sharedStrings.xml" || /xl\/worksheets\/sheet\d+\.xml/gm.test(fileName)) {
      buf = (await inflateRawAsync(entry.cfile));
      if (/xl\/worksheets\/sheet\d+\.xml/gm.test(fileName)) {
        sheets.push(buf);
      } else {
        sharedStr = buf;
      }
    }
  }
  sheetsEns = [];
  sharedJson = xml2json.toJson(sharedStr);
  sheetArr = [];
  for (m = 0, len2 = sheets.length; m < len2; m++) {
    sheetStr = sheets[m];
    ens = [];
    sheet = xml2json.toJson(sheetStr);
    if (sheet.worksheet.sheetData.row === void 0) {
      continue;
    }
    if (!isArray(sheet.worksheet.sheetData.row)) {
      sheet.worksheet.sheetData.row = [sheet.worksheet.sheetData.row];
    }
    for (i = n = 0, ref2 = sheet.worksheet.sheetData.row.length; 0 <= ref2 ? n < ref2 : n > ref2; i = 0 <= ref2 ? ++n : --n) {
      row = sheet.worksheet.sheetData.row[i];
      if (!row.c) {
        continue;
      }
      if (!isArray(row.c)) {
        row.c = [row.c];
      }
      cs = row.c;
      enr = [];
      ens[parseInt(row.r) - 1] = enr;
      numcrArr = [];
      for (o = 0, len3 = cs.length; o < len3; o++) {
        cEle = cs[o];
        crStr = cEle.r;
        crStr = crStr.replace(row.r, "");
        numcr = charToNum(crStr) - 1;
        numcrArr.push(numcr);
        if (cEle.v === void 0) {
          continue;
        }
        vStr = cEle.v["$t"];
        if (vStr === "") {
          enr[numcr] = vStr;
          continue;
        }
        if (cEle.t === "s") {
          if (!sharedJson.sst.si[vStr]) {
            continue;
          }
          if (sharedJson.sst.si[vStr].t) {
            vStr = sharedJson.sst.si[vStr].t["$t"];
          } else {
            vStr2 = "";
            if (!sharedJson.sst.si[vStr].r) {
              sharedJson.sst.si[vStr].r = [];
            } else if (!isArray(sharedJson.sst.si[vStr].r)) {
              sharedJson.sst.si[vStr].r = [sharedJson.sst.si[vStr].r];
            }
            ref3 = sharedJson.sst.si[vStr].r;
            for (p = 0, len4 = ref3.length; p < len4; p++) {
              sir = ref3[p];
              if (sir.t === void 0 || sir.t["$t"] === void 0) {
                continue;
              }
              vStr2 += sir.t["$t"];
            }
            vStr = vStr2;
          }
        }
        enr[numcr] = vStr;
      }
    }
    sheetArr.push(ens);
  }
  return sheetArr;
};

function getExcelEns(sharedStr, sheets) {
  var cEle, cont, crStr, cs, enr, ens, headsArr, i, k, l, len1, len2, len3, len4, len5, m, n, numcr, numcrArr, o, p, q, r, ref2, ref3, ref4, ref5, ref6, row, sharedJson, sheet, sheetHeadsArr, sheetStr, sheetsEns, sir, vStr, vStr2;
  sheetsEns = [];

  /*
  object: Returns a Javascript object instead of a JSON string
  reversible: Makes the JSON reversible to XML (*)
  coerce: Makes type coercion. i.e.: numbers and booleans present in attributes and element values are converted from string to its correspondent data types.
  trim: Removes leading and trailing whitespaces as well as line terminators in element values.
  sanitize: Sanitizes the following characters present in element values:
  var chars =  {  
    '<': '&lt;',
    '>': '&gt;',
    '&': '&amp;',
    '"': '&quot;',
    "'": '&apos;'
  };
    */
  if (isString(sharedStr) || Buffer.isBuffer(sharedStr)) {
    sharedJson = xml2json.toJson(sharedStr);
  } else {
    sharedJson = sharedStr;
  }
  sheetHeadsArr = [];
  for (l = 0, len1 = sheets.length; l < len1; l++) {
    sheetStr = sheets[l];
    ens = [];
    if (isString(sheetStr) || Buffer.isBuffer(sheetStr)) {
      sheet = xml2json.toJson(sheetStr);
    } else {
      sheet = sheetStr;
    }
    row = sheet.worksheet.sheetData.row;
    if (row === void 0 || row.length === void 0 || row.length < 2) {
      continue;
    }
    headsArr = [];
    sheetHeadsArr.push(headsArr);
    if (sheet.worksheet.sheetData.row[1].c === void 0) {
      sheet.worksheet.sheetData.row[1].c = [];
    } else if (!isArray(sheet.worksheet.sheetData.row[1].c)) {
      sheet.worksheet.sheetData.row[1].c = [sheet.worksheet.sheetData.row[1].c];
    }
    ref2 = sheet.worksheet.sheetData.row[1].c;
    for (m = 0, len2 = ref2.length; m < len2; m++) {
      cEle = ref2[m];
      crStr = cEle.r.toString();
      crStr = crStr.replace("2", "");
      if (cEle.v === void 0) {
        continue;
      }
      vStr = cEle.v["$t"];
      if (vStr === "") {
        enr[numcr] = vStr;
        continue;
      }
      if (!sharedJson.sst.si[vStr]) {
        continue;
      }
      if (sharedJson.sst.si[vStr].t) {
        vStr = sharedJson.sst.si[vStr].t["$t"];
      } else {
        vStr2 = "";
        if (!sharedJson.sst.si[vStr].r) {
          sharedJson.sst.si[vStr].r = [];
        } else if (!isArray(sharedJson.sst.si[vStr].r)) {
          sharedJson.sst.si[vStr].r = [sharedJson.sst.si[vStr].r];
        }
        ref3 = sharedJson.sst.si[vStr].r;
        for (n = 0, len3 = ref3.length; n < len3; n++) {
          sir = ref3[n];
          if (sir.t === void 0 || sir.t["$t"] === void 0) {
            continue;
          }
          vStr2 += sir.t["$t"];
        }
        vStr = vStr2;
      }
      numcr = charToNum(crStr);
      headsArr[numcr] = vStr;
    }
    for (i = o = 2, ref4 = sheet.worksheet.sheetData.row.length; 2 <= ref4 ? o < ref4 : o > ref4; i = 2 <= ref4 ? ++o : --o) {
      row = sheet.worksheet.sheetData.row[i];
      cs = row.c;
      if (cs === void 0 || cs === null) {
        cs = [];
      } else if (!isArray(cs)) {
        cs = [cs];
      }
      enr = {};
      ens.push(enr);
      numcrArr = [];
      for (p = 0, len4 = cs.length; p < len4; p++) {
        cEle = cs[p];
        crStr = cEle.r;
        crStr = crStr.replace(row.r, "");
        numcr = charToNum(crStr);
        numcrArr.push(numcr);
        if (cEle.v === void 0) {
          continue;
        }
        vStr = cEle.v["$t"];
        if (vStr === "") {
          enr[numcr] = vStr;
          continue;
        }
        if (cEle.t === "s") {
          if (!sharedJson.sst.si[vStr]) {
            continue;
          }
          if (sharedJson.sst.si[vStr].t) {
            vStr = sharedJson.sst.si[vStr].t["$t"];
          } else {
            vStr2 = "";
            if (!sharedJson.sst.si[vStr].r) {
              sharedJson.sst.si[vStr].r = [];
            } else if (!isArray(sharedJson.sst.si[vStr].r)) {
              sharedJson.sst.si[vStr].r = [sharedJson.sst.si[vStr].r];
            }
            ref5 = sharedJson.sst.si[vStr].r;
            for (q = 0, len5 = ref5.length; q < len5; q++) {
              sir = ref5[q];
              if (sir.t === void 0 || sir.t["$t"] === void 0) {
                continue;
              }
              vStr2 += sir.t["$t"];
            }
            vStr = vStr2;
          }
        } else {
          if (!cEle.hasOwnProperty("f") && !isNaN(Number(vStr))) {
            vStr = Number(vStr);
          }
        }
        cont = headsArr[numcr];
        if (cont !== void 0) {
          enr[cont] = vStr;
        }
      }
      for (k = r = 0, ref6 = headsArr.length; 0 <= ref6 ? r < ref6 : r > ref6; k = 0 <= ref6 ? ++r : --r) {
        if (numcrArr.indexOf(k) !== -1) {
          continue;
        }
        enr[headsArr[k]] = "";
      }
    }
    sheetsEns.push(ens);
  }
  return {
    sheetsEns: sheetsEns,
    sheetHeadsArr: sheetHeadsArr
  };
};



module.exports={
  charPlus ,
  charToNum ,
  renderPath,
  render ,
  getExcelEns ,
  renderExcel ,
  renderExcelCb ,
  getExcelArr ,
};