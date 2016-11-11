function _asyncToGenerator(fn) { return function () { var gen = fn.apply(this, arguments); return new Promise(function (resolve, reject) { function step(key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { return Promise.resolve(value).then(function (value) { return step("next", value); }, function (err) { return step("throw", err); }); } } return step("next"); }); }; }

(function () {
  var DOMParser, Hzip, Promise_fromCallback, Promise_fromStandard, Promise_sleep, charPlus, charToNum, co, crypto, drawingBuf2, drawingRelBuf2, ejs, ejs4xlx, existsAsync, fs, getExcelArr, getExcelEns, inflateRawAsync, isArray, isFunction, isObject, isString, isType, path, qr, readFileAsync, render, renderExcel, renderExcelCb, renderPath, replaceLast, sharedStrings2Prx, sheetEntrieRel2, sheetSufStr, str2Xml, wrap, xjOp, xml2json, xmldom, zlib;

  isType = function (type) {
    return function (obj) {
      return Object.prototype.toString.call(obj) === "[object " + type + "]";
    };
  };

  isObject = isType("Object");

  isString = isType("String");

  isArray = Array.isArray || isType("Array");

  isFunction = isType("Function");

  replaceLast = function (tt, what, replacement) {
    var mthArr, num;
    mthArr = tt.match(what);
    num = 0;
    return tt.replace(what, function (s) {
      num++;
      if (num === mthArr.length) {
        return replacement;
      }
      return s;
    });
  };

  fs = require("fs");

  path = require("path");

  zlib = require("zlib");

  crypto = require("crypto");

  qr = require("./lib/qr-image");

  if (typeof co === "undefined") {
    co = require("./lib/co");
  }

  ejs = require("./lib/ejs");

  Hzip = require("./lib/hzip");

  Hzip = require("./lib/hzip");

  xml2json = require("./lib/xml2json");

  xmldom = require("./lib/xmldom");

  /*in
  qr = require "qr-image"
  co = require "co" if typeof co is "undefined"
  ejs = require "ejs"
  Hzip = require "hzip"
  xml2json = require "xml2json"
  xmldom = require "xmldom"
   */

  ejs4xlx = require("./ejs4xlx");

  DOMParser = xmldom.DOMParser;

  wrap = function (func) {
    return function () {
      var rvObj;
      rvObj = func.apply(this, arguments);
      return co(rvObj);
    };
  };

  Promise_fromCallback = function (cb, t) {
    return function () {
      var args;
      args = Array.from(arguments);
      if (!t) {
        t = this;
      }
      return new Promise(function (resolve, reject) {
        args.push(function (data) {
          resolve(data);
        });
        if (cb) {
          cb.apply(t, args);
        } else {
          resolve();
        }
      });
    };
  };

  Promise_fromStandard = function (cb, t) {
    return function () {
      var args;
      args = Array.from(arguments);
      if (!t) {
        t = this;
      }
      return new Promise(function (resolve, reject) {
        args.push(function (err, data) {
          if (err) {
            reject(err);
            return;
          }
          resolve(data);
        });
        if (cb) {
          cb.apply(t, args);
        } else {
          resolve();
        }
      });
    };
  };

  Promise_sleep = function (time) {
    return new Promise(function (resolve, reject) {
      if (time > 0) {
        setTimeout(function () {
          resolve();
        }, time);
      } else {
        setImmediate(function () {
          resolve();
        });
      }
    });
  };

  existsAsync = Promise_fromCallback(fs.exists, fs);

  readFileAsync = Promise_fromStandard(fs.readFile, fs);

  inflateRawAsync = Promise_fromStandard(zlib.inflateRaw, zlib);

  render = function () {
    var _ref = _asyncToGenerator(function* (buffer, filter, _data_, hzip, options) {
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
        sharedStrings: [],
        _ss_len: 0
      };
      sharedStrings2 = [new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">")];
      data._ps_ = function (val) {
        var _ss_len, arr, index;
        val = str2Xml(val);
        arr = data._acVar_.sharedStrings;
        index = arr.indexOf(val, -200);
        _ss_len = data._acVar_._ss_len;
        if (index === -1) {
          arr.push(val);
          if (arr.length > 200) {
            arr.shift();
          }
          sharedStrings2.push(new Buffer("<si><t xml:space=\"preserve\">" + val + "</t></si>"));
          _ss_len++;
          data._acVar_._ss_len = _ss_len;
          index = _ss_len - 1;
        } else {
          index = _ss_len - (arr.length - index);
        }
        return String(index);
      };
      for (l = 0, len1 = filter.length; l < len1; l++) {
        flt = filter[l];
        if (!flt.notEjs) {
          str = ejs.parse(flt.buffer);
          anonymous = eval("(wrap(function* anonymous(_args) {\n" + str + "\n}))");
          buffer2 = yield anonymous.call(this, data);
        } else {
          buffer2 = flt.buffer;
        }
        yield updateEntryAsync.apply(hzip, [flt.path, buffer2]);
      }
      sharedStrings2.push(new Buffer("</sst>"));
      buffer2 = Buffer.concat(sharedStrings2);
      yield updateEntryAsync.apply(hzip, ["xl/sharedStrings.xml", buffer2]);
      return hzip.buffer;
    });

    return function render(_x, _x2, _x3, _x4, _x5) {
      return _ref.apply(this, arguments);
    };
  }();

  sheetSufStr = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><%\nvar _data_ = _args._data_;\nvar _charPlus_ = _args._charPlus_;\nvar _charToNum_ = _args._charToNum_;\nvar _str2Xml_ = _args._str2Xml_;\nvar _hideWorkbook_ = _args._hideWorkbook_;\nvar _ps_ = _args._ps_;\nvar _pi_ = _args._pi_;\nvar _pf_ = _args._pf_;\nvar _acVar_ = _args._acVar_;\nvar _r = 0;\nvar _c = 0;\nvar _row = 0;\nvar _col = \"\";\nvar _rc = \"\";\nvar _img_ = _args._img_;\nvar _qrcode_ = _args._qrcode_;\nvar _mergeCellArr_ = [];\nvar _mergeCellFn_ = function(mclStr) {\n	_mergeCellArr_.push(mclStr);\n};\nvar _hyperlinkArr_ = [];\nvar _outlineLevel_ = _args._outlineLevel_;\n%>");

  drawingRelBuf2 = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>");

  sheetEntrieRel2 = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>");

  drawingBuf2 = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"></xdr:wsDr>");

  sharedStrings2Prx = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">");

  xjOp = {
    object: true,
    reversible: true,
    coerce: false,
    trim: false,
    sanitize: false
  };

  renderExcelCb = function (exlBuf, _data_, callback) {
    renderExcel(exlBuf, _data_).then(function (buf2) {
      callback(null, buf2);
    })["catch"](function (err) {
      callback(err);
    });
  };

  renderExcel = function () {
    var _ref2 = _asyncToGenerator(function* (exlBuf, _data_) {
      var anonymous, attr, attr0, attr_r, begin, buffer2, cEl, cElArr, cItem, data, doc, documentElement, end, endElement, entry, hyperlink, hzip, i, idx, l, len1, len10, len11, len12, len13, len2, len3, len4, len5, len6, len7, len8, len9, location, m, m_c_i, mciNum, mciNumArr, mergeCell, mergeCellsDomEl, n, o, p, pageMarginsDomEl, phoneticPr, phoneticPrDomEl, q, r, reXmlEq, ref, ref0, ref1, ref2, ref3, ref4, ref5, ref6, ref7, ref8, ref9, refArr, row, rowEl, rowElArr, sharedStrings2, sheetBuf, sheetBuf2, sheetDataDomEl, sheetDataElementState, sheetEntrieRels, sheetEntries, sheetObj, shsEntry, shsObj, shsStr, si, si2, sirTp, startElement, str2, u, updateEntryAsync, v, w, workbookBuf, workbookEntry, x, xjOpTmp, y, z;
      data = {
        _data_: _data_
      };
      data._charPlus_ = charPlus;
      data._charToNum_ = charToNum;
      data._str2Xml_ = str2Xml;
      data._acVar_ = {
        sharedStrings: []
      };
      sharedStrings2 = [sharedStrings2Prx];
      data._ps_ = function (str, buf) {
        var _ss_len, arr, i, index, l, ref2, tmpStr, val;
        if (str === void 0) {
          str = "";
        } else if (str === null) {
          str = "NULL";
        }
        str = str.toString();
        if (str === "") {
          for (i = l = ref2 = buf.length - 1; ref2 <= -1 ? l < -1 : l > -1; i = ref2 <= -1 ? ++l : --l) {
            tmpStr = buf[i].toString();
            if (/<v>/gm.test(tmpStr)) {
              buf[i] = new Buffer(replaceLast(replaceLast(tmpStr, /<v>/gm, ""), /\s+t="s"/gm, ""));
              break;
            }
          }
          buf.push = function (puhStr) {
            var index;
            puhStr = puhStr.toString();
            index = -1;
            if (puhStr.indexOf("</v>") !== -1) {
              index = Array.prototype.push.apply(buf, [new Buffer(puhStr.replace("</v>", ""))]);
              buf.push = Array.prototype.push;
            } else {
              index = Array.prototype.push.apply(buf, [new Buffer(puhStr)]);
            }
            return index;
          };
          return "";
        }
        val = str2Xml(str);
        arr = data._acVar_.sharedStrings;
        _ss_len = data._acVar_._ss_len;
        index = arr.indexOf(val, -200);
        if (index === -1) {
          arr.push(val);
          if (arr.length > 200) {
            arr.shift();
          }
          sharedStrings2.push(new Buffer("<si><t xml:space=\"preserve\">" + val + "</t></si>"));
          _ss_len++;
          data._acVar_._ss_len = _ss_len;
          index = _ss_len - 1;
        } else {
          index = _ss_len - (arr.length - index);
        }
        return String(index);
      };
      data._pf_ = function (str, buf) {
        var i, l, m, ref2, ref3, tmpStr;
        if (str === void 0 || str === null) {
          str = "";
        }
        str = str.toString();
        str = str2Xml(str);
        for (i = l = ref2 = buf.length - 1; ref2 <= -1 ? l < -1 : l > -1; i = ref2 <= -1 ? ++l : --l) {
          tmpStr = buf[i].toString();
          if (/<v>/gm.test(tmpStr) === true) {
            buf[i] = new Buffer(replaceLast(tmpStr, /<v>/gm, "<f>"));
            break;
          }
        }
        for (i = m = ref3 = buf.length - 1; ref3 <= -1 ? m < -1 : m > -1; i = ref3 <= -1 ? ++m : --m) {
          tmpStr = buf[i].toString();
          if (/\s+t="s"/gm.test(tmpStr) === true) {
            buf[i] = new Buffer(replaceLast(tmpStr, /\s+t="s"/gm, ""));
            break;
          }
        }
        buf.push = function (puhStr) {
          var index;
          puhStr = puhStr.toString();
          if (puhStr.indexOf("</v>") !== -1) {
            index = Array.prototype.push.apply(buf, [new Buffer(puhStr.replace(/<\/v>/m, "</f>"))]);
            buf.push = Array.prototype.push;
          } else {
            index = Array.prototype.push.apply(buf, [new Buffer(puhStr)]);
          }
          return index;
        };
        return String(str);
      };
      data._pi_ = function (str, buf) {
        var i, l, ref2, tmpStr;
        if (str === true || str === false || isNaN(Number(str))) {
          return data._ps_(str, buf);
        }
        for (i = l = ref2 = buf.length - 1; ref2 <= -1 ? l < -1 : l > -1; i = ref2 <= -1 ? ++l : --l) {
          tmpStr = buf[i].toString();
          if (/\s+t="s"/gm.test(tmpStr)) {
            buf[i] = new Buffer(replaceLast(tmpStr, /\s+t="s"/gm, ""));
            break;
          }
        }
        if (str == null) {
          return "0";
        }
        str = str.toString();
        str = str2Xml(str);
        return String(str);
      };
      data._outlineLevel_ = function (str, buf) {
        var i, l, ref2, strNum, tmpStr;
        strNum = Number(str);
        if (isNaN(strNum) || strNum < 1) {
          return;
        }
        for (i = l = ref2 = buf.length - 1; ref2 <= -1 ? l < -1 : l > -1; i = ref2 <= -1 ? ++l : --l) {
          tmpStr = buf[i].toString();
          if (/<row\s/gm.test(tmpStr)) {
            buf[i] = new Buffer(replaceLast(tmpStr, /<row\s/gm, "<row outlineLevel=\"" + str + "\" "));
            break;
          }
        }
      };
      data._acVar_ = {
        sharedStrings: [],
        _ss_len: 0
      };
      hzip = new Hzip(exlBuf);
      updateEntryAsync = Promise_fromStandard(hzip.updateEntry, hzip);
      yield updateEntryAsync.apply(hzip, ["xl/calcChain.xml"]);
      sheetEntries = [];
      sheetEntrieRels = [];
      ref2 = hzip.entries;
      for (l = 0, len1 = ref2.length; l < len1; l++) {
        entry = ref2[l];
        if (entry.fileName.indexOf("xl/worksheets/sheet") === 0) {
          sheetEntries.push(entry);
        } else if (entry.fileName.indexOf("xl/worksheets/_rels/") === 0) {
          sheetEntrieRels.push(entry);
        }
      }
      sheetEntries.sort(function (arg0, arg1) {
        return arg0.fileName > arg1.fileName;
      });
      sheetEntrieRels.sort(function (arg0, arg1) {
        return arg0.fileName > arg1.fileName;
      });
      data._img_ = function () {
        var _ref3 = _asyncToGenerator(function* (imgOpt, fileName, rowNum, cellNum) {
          var cfileName, drawingBuf, drawingObj, drawingRelBuf, drawingRelObj, drawingRelStr, drawingStr, entryImgTmp, entryTmp, eny, err, hashMd5, imgBaseName, imgBuf, imgPh, itHs, len2, len3, len4, len5, m, md5Str, n, o, p, ref3, ref4, ref5, ref6, sei, xdr_frt;
          if (isString(imgOpt) || Buffer.isBuffer(imgOpt)) {
            imgOpt = {
              imgPh: imgOpt
            };
          }
          imgOpt = imgOpt || {};
          imgPh = imgOpt.imgPh;
          xdr_frt = imgOpt.xdr_frt;
          if (!imgOpt.cellNumAdd) {
            imgOpt.cellNumAdd = 1;
          }
          if (!imgOpt.rowNumAdd) {
            imgOpt.rowNumAdd = 1;
          }
          imgOpt.cellNumAdd = Number(imgOpt.cellNumAdd) - 1;
          imgOpt.rowNumAdd = Number(imgOpt.rowNumAdd) - 1;
          imgBaseName = void 0;
          imgBuf = void 0;
          if (!imgPh) {
            return "";
          }
          if (!Buffer.isBuffer(imgPh) && !isString(imgPh)) {
            return "";
          }
          if (isString(imgPh)) {
            try {
              imgBuf = yield readFileAsync(imgPh);
            } catch (error) {
              err = error;
              return "";
            }
            imgBaseName = path.basename(imgPh);
          } else {
            imgBuf = imgPh;
          }
          hashMd5 = crypto.createHash("md5");
          md5Str = hashMd5.update(imgBuf).digest("hex");
          md5Str = "a" + md5Str;
          if (!imgBaseName) {
            imgBaseName = md5Str + ".png";
          }
          cfileName = "xl/media/" + md5Str + ".png";
          itHs = false;
          ref3 = hzip.entries;
          for (m = 0, len2 = ref3.length; m < len2; m++) {
            entryTmp = ref3[m];
            if (entryTmp.fileName === cfileName) {
              itHs = true;
              break;
            }
          }
          if (!itHs) {
            yield updateEntryAsync.apply(hzip, [cfileName, imgBuf, false]);
          }
          sei = fileName.substring(fileName.length - 5, fileName.length - 4);
          sei = Number(sei) - 1;
          drawingRelBuf = void 0;
          ref4 = hzip.entries;
          for (n = 0, len3 = ref4.length; n < len3; n++) {
            entryImgTmp = ref4[n];
            if (entryImgTmp.fileName === "xl/drawings/_rels/drawing" + (sei + 1) + ".xml.rels") {
              drawingRelBuf = yield inflateRawAsync(entryImgTmp.cfile);
              break;
            }
          }
          if (drawingRelBuf === void 0) {
            console.error("Excel模板显示动态图片之前,至少需要插入一张1像素的透明的图片,以初始化");
            return "";
          }
          drawingRelObj = xml2json.toJson(drawingRelBuf, xjOp);
          if (!drawingRelObj["Relationships"]["Relationship"]) {
            drawingRelObj["Relationships"]["Relationship"] = [];
          } else if (!isArray(drawingRelObj["Relationships"]["Relationship"])) {
            drawingRelObj["Relationships"]["Relationship"] = [drawingRelObj["Relationships"]["Relationship"]];
          }
          itHs = false;
          ref5 = drawingRelObj["Relationships"]["Relationship"];
          for (o = 0, len4 = ref5.length; o < len4; o++) {
            eny = ref5[o];
            if (md5Str === eny["Id"]) {
              itHs = true;
              break;
            }
          }
          if (!itHs) {
            drawingRelObj["Relationships"]["Relationship"].push({
              "Id": md5Str,
              "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              "Target": "../media/" + md5Str + ".png"
            });
          }
          drawingRelStr = xml2json.toXml(drawingRelObj, "", {
            reSanitize: false
          });
          yield updateEntryAsync.apply(hzip, ["xl/drawings/_rels/drawing" + (sei + 1) + ".xml.rels", new Buffer('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + drawingRelStr)]);
          drawingBuf = void 0;
          ref6 = hzip.entries;
          for (p = 0, len5 = ref6.length; p < len5; p++) {
            entryImgTmp = ref6[p];
            if (entryImgTmp.fileName === "xl/drawings/drawing" + (sei + 1) + ".xml") {
              drawingBuf = yield inflateRawAsync(entryImgTmp.cfile);
              break;
            }
          }
          if (drawingBuf === void 0) {
            return "";
          }
          drawingObj = xml2json.toJson(drawingBuf, xjOp);
          if (!drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"]) {
            drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"] = [];
          } else if (!isArray(drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"])) {
            drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"] = [drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"]];
          }
          if (xdr_frt === void 0) {
            xdr_frt = {};
          }
          if (xdr_frt["from"] === void 0) {
            xdr_frt["from"] = {};
          }
          if (xdr_frt["from"]["col"] === void 0) {
            xdr_frt["from"]["col"] = String(cellNum - 1);
          }
          if (xdr_frt["from"]["colOff"] === void 0) {
            xdr_frt["from"]["colOff"] = "0";
          }
          if (xdr_frt["from"]["row"] === void 0) {
            xdr_frt["from"]["row"] = String(rowNum - 1);
          }
          if (xdr_frt["from"]["rowOff"] === void 0) {
            xdr_frt["from"]["rowOff"] = "0";
          }
          if (xdr_frt["to"] === void 0) {
            xdr_frt["to"] = {};
          }
          if (xdr_frt["to"]["col"] === void 0) {
            xdr_frt["to"]["col"] = String(cellNum + imgOpt.cellNumAdd);
          }
          if (xdr_frt["to"]["colOff"] === void 0) {
            xdr_frt["to"]["colOff"] = "0";
          }
          if (xdr_frt["to"]["row"] === void 0) {
            xdr_frt["to"]["row"] = String(rowNum + imgOpt.rowNumAdd);
          }
          if (xdr_frt["to"]["rowOff"] === void 0) {
            xdr_frt["to"]["rowOff"] = "0";
          }
          drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"].push({
            "editAs": "oneCell",
            "xdr:from": {
              "xdr:col": {
                "$t": xdr_frt["from"]["col"]
              },
              "xdr:colOff": {
                "$t": xdr_frt["from"]["colOff"]
              },
              "xdr:row": {
                "$t": xdr_frt["from"]["row"]
              },
              "xdr:rowOff": {
                "$t": xdr_frt["from"]["rowOff"]
              }
            },
            "xdr:to": {
              "xdr:col": {
                "$t": xdr_frt["to"]["col"]
              },
              "xdr:colOff": {
                "$t": xdr_frt["to"]["colOff"]
              },
              "xdr:row": {
                "$t": xdr_frt["to"]["row"]
              },
              "xdr:rowOff": {
                "$t": xdr_frt["to"]["rowOff"]
              }
            },
            "xdr:pic": {
              "xdr:nvPicPr": {
                "xdr:cNvPr": {
                  "id": "2",
                  "name": imgBaseName,
                  "descr": imgBaseName
                },
                "xdr:cNvPicPr": {
                  "a:picLocks": {
                    "noChangeAspect": "1"
                  }
                }
              },
              "xdr:blipFill": {
                "a:blip": {
                  "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                  "r:embed": md5Str
                },
                "a:stretch": {
                  "a:fillRect": {
                    "$t": ""
                  }
                }
              },
              "xdr:spPr": {
                "a:xfrm": {
                  "a:off": {
                    "x": "0",
                    "y": "0"
                  },
                  "a:ext": {
                    "cx": "2342857",
                    "cy": "1780953"
                  }
                },
                "a:prstGeom": {
                  "prst": "rect",
                  "a:avLst": {
                    "$t": ""
                  }
                }
              }
            },
            "xdr:clientData": {
              "$t": ""
            }
          });
          drawingStr = xml2json.toXml(drawingObj, "", {
            reSanitize: false
          });
          yield updateEntryAsync.apply(hzip, ["xl/drawings/drawing" + (sei + 1) + ".xml", new Buffer('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + drawingStr)]);
          return "";
        });

        return function (_x8, _x9, _x10, _x11) {
          return _ref3.apply(this, arguments);
        };
      }();
      data._qrcode_ = function () {
        var _ref4 = _asyncToGenerator(function* (imgOpt, fileName, rowNum, cellNum) {
          var imgSt, qrBufArr, rvObj;
          if (!imgOpt || !imgOpt.text) {
            return "";
          }
          if (!imgOpt.margin) {
            imgOpt.margin = 0;
          }
          qrBufArr = [];
          imgSt = qr.image(imgOpt.text, imgOpt);
          imgSt.on("data", function (dt) {
            qrBufArr.push(dt);
          });
          yield new Promise(function (resolve, reject) {
            imgSt.on("error", function (err) {
              return reject(err);
            });
            imgSt.on("end", function () {
              return resolve();
            });
          });
          imgOpt.imgPh = Buffer.concat(qrBufArr);
          rvObj = yield data._img_(imgOpt, fileName, rowNum, cellNum);
          return rvObj;
        });

        return function (_x12, _x13, _x14, _x15) {
          return _ref4.apply(this, arguments);
        };
      }();
      workbookEntry = hzip.getEntry("xl/workbook.xml");
      workbookBuf = yield inflateRawAsync(workbookEntry.cfile);
      data._hideWorkbook_ = function (strArr) {
        var doc, documentElement, len2, m, sheetEl, sheetElArr, sheetsEl;
        if (!workbookEntry) {
          return;
        }
        if (!isArray(strArr)) {
          strArr = [strArr];
        }
        doc = new DOMParser().parseFromString(workbookBuf.toString(), 'text/xml');
        documentElement = doc.documentElement;
        sheetsEl = documentElement.getElementsByTagName("sheets")[0];
        sheetElArr = sheetsEl.getElementsByTagName("sheet");
        for (m = 0, len2 = sheetElArr.length; m < len2; m++) {
          sheetEl = sheetElArr[m];
          if (strArr.indexOf(sheetEl.getAttribute("name")) === -1) {
            continue;
          }
          sheetEl.setAttribute("state", "hidden");
        }
        workbookBuf = new Buffer(doc.toString());
      };
      shsEntry = hzip.getEntry("xl/sharedStrings.xml");
      if (shsEntry === void 0) {
        return exlBuf;
      }
      shsStr = yield inflateRawAsync(shsEntry.cfile);
      shsObj = xml2json.toJson(shsStr);
      for (i = m = 0, len2 = sheetEntries.length; m < len2; i = ++m) {
        entry = sheetEntries[i];
        sheetBuf = yield inflateRawAsync(entry.cfile);
        doc = new DOMParser().parseFromString(sheetBuf.toString(), 'text/xml');
        documentElement = doc.documentElement;
        sheetDataDomEl = documentElement.getElementsByTagName("sheetData")[0];
        if (!sheetDataDomEl) {
          continue;
        }
        mergeCellsDomEl = documentElement.getElementsByTagName("mergeCells")[0];
        if (!mergeCellsDomEl) {
          mergeCellsDomEl = doc.createElement("mergeCells");
          phoneticPrDomEl = documentElement.getElementsByTagName("phoneticPr")[0];
          pageMarginsDomEl = documentElement.getElementsByTagName("pageMargins")[0];
          if (phoneticPrDomEl) {
            documentElement.insertBefore(mergeCellsDomEl, phoneticPrDomEl);
          } else if (pageMarginsDomEl) {
            documentElement.insertBefore(mergeCellsDomEl, pageMarginsDomEl);
          } else {
            documentElement.insertBefore(mergeCellsDomEl, sheetDataDomEl.nextSibling);
          }
        }
        rowElArr = sheetDataDomEl.getElementsByTagName("row");
        for (n = 0, len3 = rowElArr.length; n < len3; n++) {
          rowEl = rowElArr[n];
          if (rowEl.attributes[0] && rowEl.attributes[0].nodeName === "r") {
            continue;
          }
          attr_r = void 0;
          idx = void 0;
          ref3 = rowEl.attributes;
          for (i = o = 0, len4 = ref3.length; o < len4; i = ++o) {
            attr = ref3[i];
            if (attr.nodeName === "r") {
              attr_r = attr;
              idx = i;
              break;
            }
          }
          if (!attr_r) {
            continue;
          }
          attr0 = rowEl.attributes[0];
          rowEl.attributes[0] = attr_r;
          rowEl.attributes[idx] = attr0;
        }
        cElArr = sheetDataDomEl.getElementsByTagName("c");
        for (p = 0, len5 = cElArr.length; p < len5; p++) {
          cEl = cElArr[p];
          if (cEl.attributes[0] && cEl.attributes[0].nodeName === "r") {
            continue;
          }
          attr_r = void 0;
          idx = void 0;
          ref4 = cEl.attributes;
          for (i = q = 0, len6 = ref4.length; q < len6; i = ++q) {
            attr = ref4[i];
            if (attr.nodeName === "r") {
              attr_r = attr;
              idx = i;
              break;
            }
          }
          if (!attr_r) {
            continue;
          }
          attr0 = cEl.attributes[0];
          cEl.attributes[0] = attr_r;
          cEl.attributes[idx] = attr0;
        }
        sheetBuf = new Buffer(doc.toString());
        xjOpTmp = {
          object: true,
          reversible: true,
          coerce: false,
          trim: false,
          sanitize: true
        };
        sheetDataElementState = "";
        startElement = xml2json.toJson.startElement;
        endElement = xml2json.toJson.endElement;
        xjOpTmp.startElement = function (elementName, attrs) {
          if (elementName === "sheetData") {
            sheetDataElementState = "start";
          }
          startElement.apply(this, arguments);
        };
        xjOpTmp.endElement = function (elementName) {
          if (elementName === "sheetData") {
            sheetDataElementState = "end";
          }
          endElement.apply(this, arguments);
        };
        xjOpTmp.sanitizeFn = function (value) {
          if (!isString(value)) {
            return value;
          }
          if (sheetDataElementState === "start") {
            return value;
          }
          value = value.replace(/[&<>"']/gm, function (g1, g2) {
            if (g1 === "&") {
              return "&amp;";
            } else if (g1 === "<") {
              return "&lt;";
            } else if (g1 === ">") {
              return "&gt;";
            } else if (g1 === "\"") {
              return "&quot;";
            } else if (g1 === "'") {
              return "&apos;";
            }
            return g1;
          });
          return value;
        };
        sheetObj = xml2json.toJson(sheetBuf, xjOpTmp);
        if (sheetObj.worksheet.sheetData.row === void 0) {
          continue;
        } else if (!isArray(sheetObj.worksheet.sheetData.row)) {
          sheetObj.worksheet.sheetData.row = [sheetObj.worksheet.sheetData.row];
        }
        if (sheetObj.worksheet.mergeCells !== void 0 && sheetObj.worksheet.mergeCells.mergeCell !== void 0) {
          if (!sheetObj.worksheet.mergeCells.mergeCell) {
            sheetObj.worksheet.mergeCells.mergeCell = [];
          } else if (!isArray(sheetObj.worksheet.mergeCells.mergeCell)) {
            sheetObj.worksheet.mergeCells.mergeCell = [sheetObj.worksheet.mergeCells.mergeCell];
          }
        }
        ref5 = sheetObj.worksheet.sheetData.row;
        for (r = 0, len7 = ref5.length; r < len7; r++) {
          row = ref5[r];
          if (row.c !== void 0) {
            if (!row.c) {
              row.c = [];
            } else if (!isArray(row.c)) {
              row.c = [row.c];
            }
            ref6 = row.c;
            for (u = 0, len8 = ref6.length; u < len8; u++) {
              cItem = ref6[u];
              if (cItem.t === "s" && cItem.v && !isNaN(Number(cItem.v["$t"])) && !cItem.f) {
                if (!shsObj.sst.si) {
                  shsObj.sst.si = [];
                } else if (!isArray(shsObj.sst.si)) {
                  shsObj.sst.si = [shsObj.sst.si];
                }
                si = shsObj.sst.si[cItem.v["$t"]];
                phoneticPr = si.phoneticPr;
                si2 = {
                  t: {
                    "$t": ""
                  }
                };
                if (si.r !== void 0) {
                  if (!si.r) {
                    si.r = [];
                  } else if (!isArray(si.r)) {
                    si.r = [si.r];
                  }
                  ref7 = si.r;
                  for (v = 0, len9 = ref7.length; v < len9; v++) {
                    sirTp = ref7[v];
                    if (sirTp.t && sirTp.t["$t"]) {
                      si2.t["$t"] += sirTp.t["$t"];
                    }
                  }
                } else {
                  si2.t["$t"] = si.t["$t"];
                }
                cItem.v["$t"] = si2.t["$t"];
                if (cItem.v) {
                  if (!(cItem.v["$t"] === void 0 || cItem.v["$t"] === "")) {
                    begin = cItem.v["$t"].indexOf("<%");
                    end = cItem.v["$t"].indexOf("%>");
                    if (begin === -1 || end === -1) {
                      cItem.v["$t"] = "<%='" + cItem.v["$t"].replace(/'/gm, "\\'") + "'%>";
                    }
                  }
                }
              } else {
                if (cItem.f && cItem["v"] && cItem["v"]["$t"] && cItem["v"]["$t"].indexOf("<%") !== -1 && cItem["v"]["$t"].indexOf("%>") !== -1) {
                  delete cItem.f;
                  cItem.t = "s";
                } else {
                  if (cItem.f) {
                    if (cItem.f["$t"] !== void 0) {
                      cItem.f["$t"] = str2Xml(cItem.f["$t"]);
                    }
                    delete cItem["v"];
                  } else {
                    if (cItem.v && cItem.v["$t"]) {
                      cItem.v["$t"] = str2Xml(cItem.v["$t"]);
                    }
                  }
                }
              }
              if (sheetObj.worksheet.mergeCells !== void 0 && sheetObj.worksheet.mergeCells.mergeCell !== void 0) {
                mciNumArr = [];
                ref8 = sheetObj.worksheet.mergeCells.mergeCell;
                for (m_c_i = w = 0, len10 = ref8.length; w < len10; m_c_i = ++w) {
                  mergeCell = ref8[m_c_i];
                  if (mergeCell.ref !== void 0) {
                    refArr = mergeCell.ref.split(":");
                    ref0 = refArr[0];
                    ref1 = refArr[1];
                    if (!ref1 || !ref0) {
                      continue;
                    }
                    if (charToNum(cItem.r.replace(/\d+/, "")) >= charToNum(ref0.replace(/\d+/, "")) && Number(cItem.r.replace(/\D+/, "")) >= Number(ref0.replace(/\D+/, ""))) {
                      if (cItem.v !== void 0) {
                        if (!cItem.v["$t"]) {
                          cItem.v["$t"] = "";
                        }
                        cItem.v["$t"] += "<% _mergeCellArr_.push(_charPlus_('" + ref0.replace(/\d+/, '') + "',_c)+(" + Number(ref0.replace(/\D+/, "")) + "+_r)+':'+_charPlus_('" + ref1.replace(/\d+/, '') + "',_c)+(" + Number(ref1.replace(/\D+/, "")) + "+_r)) %>";
                      } else {
                        if (!cItem["$t"]) {
                          cItem["$t"] = "";
                        }
                        cItem["$t"] += "<% _mergeCellArr_.push(_charPlus_('" + ref0.replace(/\d+/, '') + "',_c)+(" + Number(ref0.replace(/\D+/, "")) + "+_r)+':'+_charPlus_('" + ref1.replace(/\d+/, '') + "',_c)+(" + Number(ref1.replace(/\D+/, "")) + "+_r)) %>";
                      }
                      mciNumArr.push(m_c_i);
                    }
                  }
                }
                for (x = 0, len11 = mciNumArr.length; x < len11; x++) {
                  mciNum = mciNumArr[x];
                  sheetObj.worksheet.mergeCells.mergeCell.splice(mciNum, 1);
                }
              }
              if (sheetObj.worksheet.hyperlinks && sheetObj.worksheet.hyperlinks.hyperlink) {
                mciNumArr = [];
                if (!isArray(sheetObj.worksheet.hyperlinks.hyperlink)) {
                  sheetObj.worksheet.hyperlinks.hyperlink = [sheetObj.worksheet.hyperlinks.hyperlink];
                }
                ref9 = sheetObj.worksheet.hyperlinks.hyperlink;
                for (m_c_i = y = 0, len12 = ref9.length; y < len12; m_c_i = ++y) {
                  hyperlink = ref9[m_c_i];
                  if (!hyperlink.ref) {
                    continue;
                  }
                  ref = hyperlink.ref;
                  location = hyperlink.location;
                  if (charToNum(cItem.r.replace(/\d+/, "")) >= charToNum(ref.replace(/\d+/, "")) && Number(cItem.r.replace(/\D+/, "")) >= Number(ref.replace(/\D+/, ""))) {
                    if (cItem.v != null) {
                      if (!cItem.v["$t"]) {
                        cItem.v["$t"] = "";
                      }
                      cItem.v["$t"] += "<% _hyperlinkArr_.push({ref:_charPlus_('" + ref.replace(/\d+/, '') + "',_c)+(" + Number(ref.replace(/\D+/, "")) + "+_r),location:'" + location.replace(/'/gm, "\\'").replace(/\n/gm, "\\n") + "'}) %>";
                    } else {
                      if (!cItem["$t"]) {
                        cItem["$t"] = "";
                      }
                      cItem["$t"] += "<% _hyperlinkArr_.push({ref:_charPlus_('" + ref.replace(/\d+/, '') + "',_c)+(" + Number(ref.replace(/\D+/, "")) + "+_r),location:'" + location.replace(/'/gm, "\\'").replace(/\n/gm, "\\n") + "'}) %>";
                    }
                    mciNumArr.push(m_c_i);
                  }
                }
                for (z = 0, len13 = mciNumArr.length; z < len13; z++) {
                  mciNum = mciNumArr[z];
                  sheetObj.worksheet.hyperlinks.hyperlink.splice(mciNum, 1);
                }
              }
            }
          }
        }
        if (sheetObj.worksheet.mergeCells !== void 0) {
          sheetObj.worksheet.mergeCells = {
            "$t": "<% if(_mergeCellArr_.length===0){_mergeCellArr_.push('A1:A1');} for(var m_cl=0; m_cl<_mergeCellArr_.length; m_cl++) { %><%-'<mergeCell ref=\"'+_mergeCellArr_[m_cl]+'\"/>'%><% } %>"
          };
        }
        if (sheetObj.worksheet.hyperlinks !== void 0) {
          sheetObj.worksheet.hyperlinks = {
            "$t": "<% for(var m_cl=0; m_cl<_hyperlinkArr_.length; m_cl++) { %><%-'<hyperlink ref=\"'+_hyperlinkArr_[m_cl].ref+'\" location=\"'+_hyperlinkArr_[m_cl].location+'\" display=\"\"/>'%><% } %>"
          };
        }
        sheetBuf2 = new Buffer(sheetSufStr.toString() + xml2json.toXml(sheetObj, "", {
          reSanitize: false
        }));
        reXmlEq = {
          reXmlEq: function (pixEq, jsStr, str) {
            if (pixEq === "=") {
              jsStr = jsStr.replace(/\n/gm, "\\n").replace(/\r/gm, "\\r").replace(/\t/gm, "\\t");
              jsStr = "_ps_(" + jsStr + ",buf)";
            } else if (pixEq === "~") {
              jsStr = jsStr.replace(/\n/gm, "\\n").replace(/\r/gm, "\\r").replace(/\t/gm, "\\t");
              jsStr = "_pi_(" + jsStr + ",buf)";
            } else if (pixEq === "#") {
              jsStr = jsStr.replace(/\n/gm, "\\n").replace(/\r/gm, "\\r").replace(/\t/gm, "\\t");
              jsStr = "_pf_(" + jsStr + ",buf)";
            }
            return {
              jsStr: jsStr,
              str: str
            };
          }
        };
        reXmlEq.fileName = entry.fileName;
        str2 = ejs4xlx.parse(sheetBuf2, reXmlEq);
        str2 = "(wrap(function* anonymous(_args) {" + str2 + "}))";
        anonymous = eval(str2);
        buffer2 = yield anonymous.call(this, data);
        yield updateEntryAsync.apply(hzip, [entry.fileName, buffer2]);
      }
      sharedStrings2.push(new Buffer("</sst>"));
      buffer2 = Buffer.concat(sharedStrings2);
      sharedStrings2 = void 0;
      yield updateEntryAsync.apply(hzip, ["xl/sharedStrings.xml", buffer2]);
      yield updateEntryAsync.apply(hzip, ["xl/workbook.xml", workbookBuf]);
      return hzip.buffer;
    });

    return function renderExcel(_x6, _x7) {
      return _ref2.apply(this, arguments);
    };
  }();

  renderPath = function () {
    var _ref5 = _asyncToGenerator(function* (ejsDir, data) {
      var buffer, config, configPath, exists, exlBuf, extname, filter, ftObj, key, l, len1, obj, val;
      configPath = ejsDir + "/config.json";
      exists = yield existsAsync(configPath);
      if (exists === false) {
        extname = path.extname(ejsDir);
        if (extname === "") {
          ejsDir = ejsDir + ".xlsx";
        }
        exlBuf = yield readFileAsync(ejsDir);
        return yield renderExcel(exlBuf, data);
      }
      config = yield readFileAsync(configPath, "utf8");
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
        ftObj.buffer = yield readFileAsync(ejsDir + "/" + key);
        filter.push(ftObj);
      }
      buffer = yield readFileAsync(ejsDir + "/" + path.basename(ejsDir) + ".zip");
      return yield render(buffer, filter, data);
    });

    return function renderPath(_x16, _x17) {
      return _ref5.apply(this, arguments);
    };
  }();

  getExcelArr = function () {
    var _ref6 = _asyncToGenerator(function* (buffer) {
      var buf, cEle, crStr, cs, enr, ens, entries, entry, fileName, hzip, i, l, len1, len2, len3, len4, m, n, numcr, numcrArr, o, p, ref2, ref3, row, sharedJson, sharedStr, sheet, sheetArr, sheetStr, sheets, sheetsEns, sir, vStr, vStr2;
      sharedStr = null;
      sheets = [];
      hzip = new Hzip(buffer);
      entries = hzip.entries;
      for (l = 0, len1 = entries.length; l < len1; l++) {
        entry = entries[l];
        fileName = entry.fileName;
        if (fileName === "xl/sharedStrings.xml" || /xl\/worksheets\/sheet\d+\.xml/gm.test(fileName)) {
          buf = yield inflateRawAsync(entry.cfile);
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
    });

    return function getExcelArr(_x18) {
      return _ref6.apply(this, arguments);
    };
  }();

  getExcelEns = function (sharedStr, sheets) {
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

  str2Xml = function (str) {
    var charTmp, i, l, ref2, s, str2;
    if (!isString(str)) {
      return str;
    }

    /*
    arr2 = []
    buf = new Buffer str
    for i in [0...buf.length]
      code = buf.readInt8 i
      continue if code is 0
      arr2.push code
    str = new Buffer(arr2).toString()
    '&lt;'  :'<',
     '&gt;'  :'>',
     '&\#40;' :'(',
     '&\#41;' :')',
     '&\#35;' :'#',
     '&quot;':'"',
     '&apos;':"'",
     '&amp;' :'&'
     */
    str2 = "";
    for (i = l = 0, ref2 = str.length; 0 <= ref2 ? l < ref2 : l > ref2; i = 0 <= ref2 ? ++l : --l) {
      charTmp = str.charCodeAt(i);
      s = str.charAt(i);
      if (charTmp <= 31 && charTmp !== 9 && charTmp !== 10 || charTmp === 127) {
        s = JSON.stringify(s);
        s = s.substring(1, s.length - 1);
        s = s.replace("\\u", "_x") + "_";
        str2 += s;
        continue;
      }
      if (s === "&") {
        s = "&amp;";
      } else if (s === "<") {
        s = "&lt;";
      } else if (s === ">") {
        s = "&gt;";
      } else if (s === "\"") {
        s = "&quot;";
      } else if (s === "'") {
        s = "&apos;";
      }
      str2 += s;
    }
    return str2;
  };

  charPlus = function (str, num) {
    var ch, i, strNum, temp;
    strNum = charToNum(str);
    strNum += num;
    if (strNum <= 0) {
      return "A";
    }
    temp = "";
    ch = "";
    while (strNum >= 1) {
      i = strNum % 26;
      if (i !== 0) {
        ch = String.fromCharCode(65 + i - 1);
        temp = ch + temp;
      } else {
        ch = "Z";
        temp = ch + temp;
        strNum--;
      }
      strNum = Math.floor(strNum / 26);
    }
    return temp;
  };

  charToNum = function (str) {
    var i, j, l, len, ref2, temp, val;
    str = new String(str);
    val = 0;
    len = str.length;
    for (j = l = 0, ref2 = len; 0 <= ref2 ? l < ref2 : l > ref2; j = 0 <= ref2 ? ++l : --l) {
      i = len - 1 - j;
      temp = str.charCodeAt(i) - 65 + 1;
      val += temp * Math.pow(26, j);
    }
    return val;
  };

  exports.charPlus = charPlus;

  exports.charToNum = charToNum;

  exports.renderPath = renderPath;

  exports.render = render;

  exports.getExcelEns = getExcelEns;

  exports.renderExcel = renderExcel;

  exports.renderExcelCb = renderExcelCb;

  exports.getExcelArr = getExcelArr;
}).call(this);