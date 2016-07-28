(function() {
  var Binding, DOMParser, Hzip, Task, Wind, charPlus, charToNum, crypto, drawingBuf2, drawingRelBuf2, ejs, ejs4xlx, err, existsAsync, fs, getExcelArr, getExcelArrCb, getExcelEns, inflateRawAsync, isArray, isFunction, isObject, isString, isType, path, readFileAsync, render, renderExcel, renderExcelCb, renderPath, replaceLast, sheetEntrieRel2, sheetSufStr, str2Xml, xjOp, xml2json, xmldom, zlib;

  isType = function(type) {
    return function(obj) {
      return Object.prototype.toString.call(obj) === "[object " + type + "]";
    };
  };

  isObject = isType("Object");

  isString = isType("String");

  isArray = Array.isArray || isType("Array");

  isFunction = isType("Function");

  replaceLast = function(tt, what, replacement) {
    var mthArr, num;
    mthArr = tt.match(what);
    num = 0;
    return tt.replace(what, function(s) {
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

  if (typeof Wind === "undefined") {
    try {
      Wind = require("./lib/Wind");
    } catch (_error) {
      err = _error;
      Wind = require("wind");
    }
  }

  ejs4xlx = void 0;

  try {
    ejs4xlx = require("./ejs4xlx");
  } catch (_error) {
    err = _error;
    ejs4xlx = require("ejs4xlx");
  }

  ejs = void 0;

  try {
    ejs = require("./lib/ejs");
  } catch (_error) {
    err = _error;
    ejs = require("ejs");
  }

  Hzip = void 0;

  try {
    Hzip = require("./lib/hzip");
  } catch (_error) {
    err = _error;
    Hzip = require("hzip");
  }

  xml2json = void 0;

  try {
    xml2json = require("./lib/xml2json");
  } catch (_error) {
    err = _error;
    xml2json = require("xml2json");
  }

  xmldom = void 0;

  try {
    xmldom = require("./lib/xmldom");
  } catch (_error) {
    err = _error;
    xmldom = require("xmldom");
  }

  DOMParser = xmldom.DOMParser;

  Task = Wind.Async.Task;

  Binding = Wind.Async.Binding;

  existsAsync = Binding.fromCallback(fs.exists);

  readFileAsync = Binding.fromStandard(fs.readFile);

  inflateRawAsync = Binding.fromStandard(zlib.inflateRaw);

  render = (function (buffer, filter, _data_, hzip, options) {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            var anonymous, buffer2, data, entries, flt, l, len1, sharedStrings2, str, updateEntryAsync;
            if (hzip === void 0 || hzip === null) {
                hzip = new Hzip(buffer);
            }
            updateEntryAsync = Binding.fromStandard(hzip.updateEntry);
            entries = hzip.entries;
            data = {
                "_data_": _data_
            };
            data._charPlus_ = charPlus;
            data._charToNum_ = charToNum;
            data._str2Xml_ = str2Xml;
            data._acVar_ = {
                "sharedStrings": [],
                "_ss_len": 0
            };
            sharedStrings2 = [new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">")];
            data._ps_ = function (val) {
                var _ss_len, arr, index;
                val = str2Xml(val);
                arr = data._acVar_.sharedStrings;
                index = arr.indexOf(val, - 200);
                _ss_len = data._acVar_._ss_len;
                if (index === - 1) {
                    arr.push(val);
                    if (arr.length > 200) {
                        arr.shift();
                    }
                    sharedStrings2.push(new Buffer("<si><t xml:space=\"preserve\">" + val + "</t></si>"));
                    _ss_len ++;
                    data._acVar_._ss_len = _ss_len;
                    index = _ss_len - 1;
                } else {
                    index = _ss_len - (arr.length - index);
                }
                return String(index);
            };
            return _builder_$0.f(
                _builder_$0.e(function() {
                    (l = 0, len1 = filter.length)
                    return _builder_$0.a(function() {
                        return l < len1;
                    }, function() {
                        l ++;
                    },
                        _builder_$0.e(function() {
                            flt = filter[l];
                            return _builder_$0.f(
                                _builder_$0.e(function() {
                                    if (! flt.notEjs) {
                                        str = ejs.parse(flt.buffer);
                                        anonymous = eval(Wind.compile("async", "function anonymous(_args) {\n" + str + "\n}"));
                                        return _builder_$0.n(anonymous.call(this, data), function (_result_$) {
                                            buffer2 = _result_$;
                                            return _builder_$0.h();
                                        });
                                    } else {
                                        buffer2 = flt.buffer;
                                        return _builder_$0.h();
                                    }
                                }),
                                _builder_$0.e(function() {
                                    return _builder_$0.n(updateEntryAsync.apply(hzip, [flt.path, buffer2]), function () {
                                        return _builder_$0.h();
                                    });
                                })
                            );
                        })
                    );
                }),
                _builder_$0.e(function() {
                    sharedStrings2.push(new Buffer("</sst>"));
                    buffer2 = Buffer.concat(sharedStrings2);
                    return _builder_$0.n(updateEntryAsync.apply(hzip, ["xl/sharedStrings.xml", buffer2]), function () {
                        return _builder_$0.g(hzip.buffer);
                    });
                })
            );
        })
    );
});

  sheetSufStr = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><%\nvar _data_ = _args._data_;\nvar _charPlus_ = _args._charPlus_;\nvar _charToNum_ = _args._charToNum_;\nvar _str2Xml_ = _args._str2Xml_;\nvar _ps_ = _args._ps_;\nvar _pi_ = _args._pi_;\nvar _pf_ = _args._pf_;\nvar _acVar_ = _args._acVar_;\nvar _r = 0;\nvar _c = 0;\nvar _row = 0;\nvar _col = \"\";\nvar _rc = \"\";\nvar _imgAsync_ = _args._imgAsync_;\nvar _img_ = _args._img_;\nvar _mergeCellArr_ = [];\nvar _mergeCellFn_ = function(mclStr) {\n	_mergeCellArr_.push(mclStr);\n};\nvar _hyperlinkArr_ = [];\nvar _outlineLevel_ = _args._outlineLevel_;\n%>");

  drawingRelBuf2 = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>");

  sheetEntrieRel2 = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>");

  drawingBuf2 = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"></xdr:wsDr>");

  xjOp = {
    object: true,
    reversible: true,
    coerce: false,
    trim: false,
    sanitize: false
  };

  renderExcelCb = function(exlBuf, _data_, callback) {
    var tmpFn;
    tmpFn = (function () {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            var buf2;
            return _builder_$0.f(
                _builder_$0.l(
                    _builder_$0.e(function() {
                        return _builder_$0.n(renderExcel(exlBuf, _data_), function (_result_$) {
                            buf2 = _result_$;
                            return _builder_$0.h();
                        });
                    }),
                    function (_error) {
                        err = _error;
                        callback(err);
                        return _builder_$0.g();
                    },
                    null
                ),
                _builder_$0.e(function() {
                    callback(null, buf2);
                    return _builder_$0.h();
                })
            );
        })
    );
});
    tmpFn().start();
  };

  renderExcel = (function (exlBuf, _data_) {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            var _imgFn_, anonymous, begin, buffer2, cItem, data, doc, documentElement, end, endElement, entry, hyperlink, hzip, i, imgTk, imgTkArr, l, len1, len10, len2, len3, len4, len5, len6, len7, len8, len9, location, m, m_c_i, mciNum, mciNumArr, mergeCell, mergeCellsDomEl, n, o, p, pageMarginsDomEl, phoneticPr, phoneticPrDomEl, q, r, reXmlEq, ref, ref0, ref1, ref2, ref3, ref4, ref5, ref6, ref7, refArr, row, sharedStrings2, sheetBuf, sheetBuf2, sheetDataDomEl, sheetDataElementState, sheetEntrieRels, sheetEntries, sheetObj, shsEntry, shsObj, shsStr, si, si2, sirTp, startElement, str2, t, u, updateEntryAsync, v, xjOpTmp;
            data = {
                "_data_": _data_
            };
            data._charPlus_ = charPlus;
            data._charToNum_ = charToNum;
            data._str2Xml_ = str2Xml;
            data._acVar_ = {
                "sharedStrings": []
            };
            sharedStrings2 = [new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">")];
            data._ps_ = function (str, buf) {
                var _ss_len, arr, i, index, l, ref2, tmpStr, val;
                if (str === void 0) {
                    str = "";
                } else if (str === null) {
                    str = "NULL";
                }
                str = str.toString();
                if (str === "") {
                    for (i = l = ref2 = buf.length - 1; (ref2 <= - 1) ? (l < - 1) : (l > - 1); i = (ref2 <= - 1) ? (++ l) : (-- l)) {
                        tmpStr = buf[i].toString();
                        if (/<v>/gm.test(tmpStr)) {
                            buf[i] = new Buffer(replaceLast(replaceLast(tmpStr, /<v>/gm, ""), /\s+t="s"/gm, ""));
                            break;
                        }
                    }
                    buf.push = function (puhStr) {
                        var index;
                        puhStr = puhStr.toString();
                        index = - 1;
                        if (puhStr.indexOf("</v>") !== - 1) {
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
                index = arr.indexOf(val, - 200);
                if (index === - 1) {
                    arr.push(val);
                    if (arr.length > 200) {
                        arr.shift();
                    }
                    sharedStrings2.push(new Buffer("<si><t xml:space=\"preserve\">" + val + "</t></si>"));
                    _ss_len ++;
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
                for (i = l = ref2 = buf.length - 1; (ref2 <= - 1) ? (l < - 1) : (l > - 1); i = (ref2 <= - 1) ? (++ l) : (-- l)) {
                    tmpStr = buf[i].toString();
                    if (/<v>/gm.test(tmpStr) === true) {
                        buf[i] = new Buffer(replaceLast(tmpStr, /<v>/gm, "<f>"));
                        break;
                    }
                }
                for (i = m = ref3 = buf.length - 1; (ref3 <= - 1) ? (m < - 1) : (m > - 1); i = (ref3 <= - 1) ? (++ m) : (-- m)) {
                    tmpStr = buf[i].toString();
                    if (/\s+t="s"/gm.test(tmpStr) === true) {
                        buf[i] = new Buffer(replaceLast(tmpStr, /\s+t="s"/gm, ""));
                        break;
                    }
                }
                buf.push = function (puhStr) {
                    var index;
                    puhStr = puhStr.toString();
                    if (puhStr.indexOf("</v>") !== - 1) {
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
                if (isNaN(Number(str))) {
                    return data._ps_(str, buf);
                }
                for (i = l = ref2 = buf.length - 1; (ref2 <= - 1) ? (l < - 1) : (l > - 1); i = (ref2 <= - 1) ? (++ l) : (-- l)) {
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
                for (i = l = ref2 = buf.length - 1; (ref2 <= - 1) ? (l < - 1) : (l > - 1); i = (ref2 <= - 1) ? (++ l) : (-- l)) {
                    tmpStr = buf[i].toString();
                    if (/<row\s/gm.test(tmpStr)) {
                        buf[i] = new Buffer(replaceLast(tmpStr, /<row\s/gm, "<row outlineLevel=\"" + str + "\" "));
                        break;
                    }
                }
            };
            data._acVar_ = {
                "sharedStrings": [],
                "_ss_len": 0
            };
            hzip = new Hzip(exlBuf);
            updateEntryAsync = Binding.fromStandard(hzip.updateEntry);
            return _builder_$0.n(updateEntryAsync.apply(hzip, ["xl/calcChain.xml"]), function () {
                sheetEntries = [];
                sheetEntrieRels = [];
                ref2 = hzip.entries;
                for ((l = 0, len1 = ref2.length); l < len1; l ++) {
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
                data._img_ = (function (imgOpt, fileName, rowNum, cellNum) {
                    var _builder_$1 = Wind.b["async"];
                    var _arguments_$ = arguments;
                    return _builder_$1.m(this,
                        _builder_$1.e(function() {
                            var cfileName, drawingBuf, drawingObj, drawingRelBuf, drawingRelObj, drawingRelStr, drawingStr, entryImgTmp, entryTmp, eny, hashMd5, imgBaseName, imgBuf, imgPh, itHs, len2, len3, len4, len5, m, md5Str, n, o, p, ref3, ref4, ref5, ref6, sei, xdr_frt;
                            if (isString(imgOpt) || Buffer.isBuffer(imgOpt)) {
                                imgOpt = {
                                    "imgPh": imgOpt
                                };
                            }
                            imgOpt = imgOpt || { };
                            imgPh = imgOpt.imgPh;
                            xdr_frt = imgOpt.xdr_frt;
                            if (! imgOpt.cellNumAdd) {
                                imgOpt.cellNumAdd = 0;
                            }
                            if (! imgOpt.rowNumAdd) {
                                imgOpt.rowNumAdd = 0;
                            }
                            imgOpt.cellNumAdd = Number(imgOpt.cellNumAdd);
                            imgOpt.rowNumAdd = Number(imgOpt.rowNumAdd);
                            imgBaseName = void 0;
                            imgBuf = void 0;
                            if (! imgPh) {
                                return _builder_$1.g("");
                            }
                            if (! Buffer.isBuffer(imgPh) && ! isString(imgPh)) {
                                return _builder_$1.g("");
                            }
                            return _builder_$1.f(
                                _builder_$1.e(function() {
                                    if (isString(imgPh)) {
                                        return _builder_$1.f(
                                            _builder_$1.l(
                                                _builder_$1.e(function() {
                                                    return _builder_$1.n(readFileAsync(imgPh), function (_result_$) {
                                                        imgBuf = _result_$;
                                                        return _builder_$1.h();
                                                    });
                                                }),
                                                function (_error) {
                                                    err = _error;
                                                    return _builder_$1.g("");
                                                },
                                                null
                                            ),
                                            _builder_$1.e(function() {
                                                imgBaseName = path.basename(imgPh);
                                                return _builder_$1.h();
                                            })
                                        );
                                    } else {
                                        imgBuf = imgPh;
                                        return _builder_$1.h();
                                    }
                                }),
                                _builder_$1.e(function() {
                                    hashMd5 = crypto.createHash("md5");
                                    md5Str = hashMd5.update(imgBuf).digest("hex");
                                    md5Str = "a" + md5Str;
                                    if (! imgBaseName) {
                                        imgBaseName = md5Str + ".png";
                                    }
                                    cfileName = "xl/media/" + md5Str + ".png";
                                    itHs = false;
                                    ref3 = hzip.entries;
                                    for ((m = 0, len2 = ref3.length); m < len2; m ++) {
                                        entryTmp = ref3[m];
                                        if (entryTmp.fileName === cfileName) {
                                            itHs = true;
                                            break;
                                        }
                                    }
                                    return _builder_$1.f(
                                        _builder_$1.e(function() {
                                            if (! itHs) {
                                                return _builder_$1.n(updateEntryAsync.apply(hzip, [cfileName, imgBuf, false]), function () {
                                                    return _builder_$1.h();
                                                });
                                            } else {
                                                return _builder_$1.h();
                                            }
                                        }),
                                        _builder_$1.e(function() {
                                            sei = fileName.substring(fileName.length - 5, fileName.length - 4);
                                            sei = Number(sei) - 1;
                                            drawingRelBuf = void 0;
                                            ref4 = hzip.entries;
                                            return _builder_$1.f(
                                                _builder_$1.e(function() {
                                                    (n = 0, len3 = ref4.length)
                                                    return _builder_$1.a(function() {
                                                        return n < len3;
                                                    }, function() {
                                                        n ++;
                                                    },
                                                        _builder_$1.e(function() {
                                                            entryImgTmp = ref4[n];
                                                            if (entryImgTmp.fileName === "xl/drawings/_rels/drawing" + (sei + 1) + ".xml.rels") {
                                                                return _builder_$1.n(inflateRawAsync(entryImgTmp.cfile), function (_result_$) {
                                                                    drawingRelBuf = _result_$;
                                                                    return _builder_$1.i();
                                                                });
                                                            } else {
                                                                return _builder_$1.h();
                                                            }
                                                        })
                                                    );
                                                }),
                                                _builder_$1.e(function() {
                                                    if (drawingRelBuf === void 0) {
                                                        console.error("Excel模板显示动态图片之前,至少需要插入一张1像素的透明的图片,以初始化");
                                                        return _builder_$1.g("");
                                                    }
                                                    drawingRelObj = xml2json.toJson(drawingRelBuf, xjOp);
                                                    if (! drawingRelObj["Relationships"]["Relationship"]) {
                                                        drawingRelObj["Relationships"]["Relationship"] = [];
                                                    } else if (! isArray(drawingRelObj["Relationships"]["Relationship"])) {
                                                        drawingRelObj["Relationships"]["Relationship"] = [drawingRelObj["Relationships"]["Relationship"]];
                                                    }
                                                    itHs = false;
                                                    ref5 = drawingRelObj["Relationships"]["Relationship"];
                                                    for ((o = 0, len4 = ref5.length); o < len4; o ++) {
                                                        eny = ref5[o];
                                                        if (md5Str === eny["Id"]) {
                                                            itHs = true;
                                                            break;
                                                        }
                                                    }
                                                    if (! itHs) {
                                                        drawingRelObj["Relationships"]["Relationship"].push({
                                                            "Id": md5Str,
                                                            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                                                            "Target": "../media/" + md5Str + ".png"
                                                        });
                                                    }
                                                    drawingRelStr = xml2json.toXml(drawingRelObj, "", {
                                                        "reSanitize": false
                                                    });
                                                    return _builder_$1.n(updateEntryAsync.apply(hzip, ["xl/drawings/_rels/drawing" + (sei + 1) + ".xml.rels", new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + drawingRelStr)]), function () {
                                                        drawingBuf = void 0;
                                                        ref6 = hzip.entries;
                                                        return _builder_$1.f(
                                                            _builder_$1.e(function() {
                                                                (p = 0, len5 = ref6.length)
                                                                return _builder_$1.a(function() {
                                                                    return p < len5;
                                                                }, function() {
                                                                    p ++;
                                                                },
                                                                    _builder_$1.e(function() {
                                                                        entryImgTmp = ref6[p];
                                                                        if (entryImgTmp.fileName === "xl/drawings/drawing" + (sei + 1) + ".xml") {
                                                                            return _builder_$1.n(inflateRawAsync(entryImgTmp.cfile), function (_result_$) {
                                                                                drawingBuf = _result_$;
                                                                                return _builder_$1.i();
                                                                            });
                                                                        } else {
                                                                            return _builder_$1.h();
                                                                        }
                                                                    })
                                                                );
                                                            }),
                                                            _builder_$1.e(function() {
                                                                if (drawingBuf === void 0) {
                                                                    return _builder_$1.g("");
                                                                }
                                                                drawingObj = xml2json.toJson(drawingBuf, xjOp);
                                                                if (! drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"]) {
                                                                    drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"] = [];
                                                                } else if (! isArray(drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"])) {
                                                                    drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"] = [drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"]];
                                                                }
                                                                if (xdr_frt === void 0) {
                                                                    xdr_frt = { };
                                                                }
                                                                if (xdr_frt["from"] === void 0) {
                                                                    xdr_frt["from"] = { };
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
                                                                    xdr_frt["to"] = { };
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
                                                                    "reSanitize": false
                                                                });
                                                                return _builder_$1.n(updateEntryAsync.apply(hzip, ["xl/drawings/drawing" + (sei + 1) + ".xml", new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + drawingStr)]), function () {
                                                                    return _builder_$1.g("");
                                                                });
                                                            })
                                                        );
                                                    });
                                                })
                                            );
                                        })
                                    );
                                })
                            );
                        })
                    );
                });
                imgTkArr = [];
                _imgFn_ = data._img_;
                data._img_ = function () {
                    var tk;
                    tk = _imgFn_.apply(this, arguments);
                    imgTkArr.push(tk);
                    return "";
                };
                shsEntry = hzip.getEntry("xl/sharedStrings.xml");
                if (shsEntry === void 0) {
                    return _builder_$0.g(exlBuf);
                }
                return _builder_$0.n(inflateRawAsync(shsEntry.cfile), function (_result_$) {
                    shsStr = _result_$;
                    shsObj = xml2json.toJson(shsStr);
                    return _builder_$0.f(
                        _builder_$0.e(function() {
                            (i = m = 0, len2 = sheetEntries.length)
                            return _builder_$0.a(function() {
                                return m < len2;
                            }, function() {
                                i = ++ m;
                            },
                                _builder_$0.e(function() {
                                    entry = sheetEntries[i];
                                    return _builder_$0.n(inflateRawAsync(entry.cfile), function (_result_$) {
                                        sheetBuf = _result_$;
                                        doc = (new DOMParser()).parseFromString(sheetBuf.toString(), "text/xml");
                                        documentElement = doc.documentElement;
                                        sheetDataDomEl = documentElement.getElementsByTagName("sheetData")[0];
                                        if (! sheetDataDomEl) {
                                            return _builder_$0.j();
                                        }
                                        mergeCellsDomEl = documentElement.getElementsByTagName("mergeCells")[0];
                                        if (! mergeCellsDomEl) {
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
                                            sheetBuf = new Buffer(doc.toString());
                                        }
                                        xjOpTmp = {
                                            "object": true,
                                            "reversible": true,
                                            "coerce": false,
                                            "trim": false,
                                            "sanitize": true
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
                                            if (! isString(value)) {
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
                                            return _builder_$0.j();
                                        } else if (! isArray(sheetObj.worksheet.sheetData.row)) {
                                            sheetObj.worksheet.sheetData.row = [sheetObj.worksheet.sheetData.row];
                                        }
                                        if (sheetObj.worksheet.mergeCells !== void 0 && sheetObj.worksheet.mergeCells.mergeCell !== void 0) {
                                            if (! sheetObj.worksheet.mergeCells.mergeCell) {
                                                sheetObj.worksheet.mergeCells.mergeCell = [];
                                            } else if (! isArray(sheetObj.worksheet.mergeCells.mergeCell)) {
                                                sheetObj.worksheet.mergeCells.mergeCell = [sheetObj.worksheet.mergeCells.mergeCell];
                                            }
                                        }
                                        ref3 = sheetObj.worksheet.sheetData.row;
                                        for ((n = 0, len3 = ref3.length); n < len3; n ++) {
                                            row = ref3[n];
                                            if (row.c !== void 0) {
                                                if (! row.c) {
                                                    row.c = [];
                                                } else if (! isArray(row.c)) {
                                                    row.c = [row.c];
                                                }
                                                ref4 = row.c;
                                                for ((o = 0, len4 = ref4.length); o < len4; o ++) {
                                                    cItem = ref4[o];
                                                    if (cItem.t === "s" && cItem.v && ! isNaN(Number(cItem.v["$t"])) && ! cItem.f) {
                                                        if (! shsObj.sst.si) {
                                                            shsObj.sst.si = [];
                                                        } else if (! isArray(shsObj.sst.si)) {
                                                            shsObj.sst.si = [shsObj.sst.si];
                                                        }
                                                        si = shsObj.sst.si[cItem.v["$t"]];
                                                        phoneticPr = si.phoneticPr;
                                                        si2 = {
                                                            "t": {
                                                                "$t": ""
                                                            }
                                                        };
                                                        if (si.r !== void 0) {
                                                            if (! si.r) {
                                                                si.r = [];
                                                            } else if (! isArray(si.r)) {
                                                                si.r = [si.r];
                                                            }
                                                            ref5 = si.r;
                                                            for ((p = 0, len5 = ref5.length); p < len5; p ++) {
                                                                sirTp = ref5[p];
                                                                if (sirTp.t && sirTp.t["$t"]) {
                                                                    si2.t["$t"] += sirTp.t["$t"];
                                                                }
                                                            }
                                                        } else {
                                                            si2.t["$t"] = si.t["$t"];
                                                        }
                                                        cItem.v["$t"] = si2.t["$t"];
                                                        if (cItem.v) {
                                                            if (! (cItem.v["$t"] === void 0 || cItem.v["$t"] === "")) {
                                                                begin = cItem.v["$t"].indexOf("<%");
                                                                end = cItem.v["$t"].indexOf("%>");
                                                                if (begin === - 1 || end === - 1) {
                                                                    cItem.v["$t"] = "<%='" + cItem.v["$t"].replace(/'/gm, "\\'") + "'%>";
                                                                }
                                                            }
                                                        }
                                                    } else {
                                                        if (cItem.f && cItem["v"] && cItem["v"]["$t"] && cItem["v"]["$t"].indexOf("<%") !== - 1 && cItem["v"]["$t"].indexOf("%>") !== - 1) {
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
                                                        ref6 = sheetObj.worksheet.mergeCells.mergeCell;
                                                        for ((m_c_i = q = 0, len6 = ref6.length); q < len6; m_c_i = ++ q) {
                                                            mergeCell = ref6[m_c_i];
                                                            if (mergeCell.ref !== void 0) {
                                                                refArr = mergeCell.ref.split(":");
                                                                ref0 = refArr[0];
                                                                ref1 = refArr[1];
                                                                if (! ref1 || ! ref0) {
                                                                    continue;
                                                                }
                                                                if (charToNum(cItem.r.replace(/\d+/, "")) >= charToNum(ref0.replace(/\d+/, "")) && Number(cItem.r.replace(/\D+/, "")) >= Number(ref0.replace(/\D+/, ""))) {
                                                                    if (cItem.v !== void 0) {
                                                                        if (! cItem.v["$t"]) {
                                                                            cItem.v["$t"] = "";
                                                                        }
                                                                        cItem.v["$t"] += "<% _mergeCellArr_.push(_charPlus_('" + ref0.replace(/\d+/, "") + "',_c)+(" + Number(ref0.replace(/\D+/, "")) + "+_r)+':'+_charPlus_('" + ref1.replace(/\d+/, "") + "',_c)+(" + Number(ref1.replace(/\D+/, "")) + "+_r)) %>";
                                                                    } else {
                                                                        if (! cItem["$t"]) {
                                                                            cItem["$t"] = "";
                                                                        }
                                                                        cItem["$t"] += "<% _mergeCellArr_.push(_charPlus_('" + ref0.replace(/\d+/, "") + "',_c)+(" + Number(ref0.replace(/\D+/, "")) + "+_r)+':'+_charPlus_('" + ref1.replace(/\d+/, "") + "',_c)+(" + Number(ref1.replace(/\D+/, "")) + "+_r)) %>";
                                                                    }
                                                                    mciNumArr.push(m_c_i);
                                                                }
                                                            }
                                                        }
                                                        for ((r = 0, len7 = mciNumArr.length); r < len7; r ++) {
                                                            mciNum = mciNumArr[r];
                                                            sheetObj.worksheet.mergeCells.mergeCell.splice(mciNum, 1);
                                                        }
                                                    }
                                                    if (sheetObj.worksheet.hyperlinks && sheetObj.worksheet.hyperlinks.hyperlink) {
                                                        mciNumArr = [];
                                                        if (! isArray(sheetObj.worksheet.hyperlinks.hyperlink)) {
                                                            sheetObj.worksheet.hyperlinks.hyperlink = [sheetObj.worksheet.hyperlinks.hyperlink];
                                                        }
                                                        ref7 = sheetObj.worksheet.hyperlinks.hyperlink;
                                                        for ((m_c_i = t = 0, len8 = ref7.length); t < len8; m_c_i = ++ t) {
                                                            hyperlink = ref7[m_c_i];
                                                            if (! hyperlink.ref) {
                                                                continue;
                                                            }
                                                            ref = hyperlink.ref;
                                                            location = hyperlink.location;
                                                            if (charToNum(cItem.r.replace(/\d+/, "")) >= charToNum(ref.replace(/\d+/, "")) && Number(cItem.r.replace(/\D+/, "")) >= Number(ref.replace(/\D+/, ""))) {
                                                                if (cItem.v != null) {
                                                                    if (! cItem.v["$t"]) {
                                                                        cItem.v["$t"] = "";
                                                                    }
                                                                    cItem.v["$t"] += "<% _hyperlinkArr_.push({ref:_charPlus_('" + ref.replace(/\d+/, "") + "',_c)+(" + Number(ref.replace(/\D+/, "")) + "+_r),location:'" + location.replace(/'/gm, "\\'").replace(/\n/gm, "\\n") + "'}) %>";
                                                                } else {
                                                                    if (! cItem["$t"]) {
                                                                        cItem["$t"] = "";
                                                                    }
                                                                    cItem["$t"] += "<% _hyperlinkArr_.push({ref:_charPlus_('" + ref.replace(/\d+/, "") + "',_c)+(" + Number(ref.replace(/\D+/, "")) + "+_r),location:'" + location.replace(/'/gm, "\\'").replace(/\n/gm, "\\n") + "'}) %>";
                                                                }
                                                                mciNumArr.push(m_c_i);
                                                            }
                                                        }
                                                        for ((u = 0, len9 = mciNumArr.length); u < len9; u ++) {
                                                            mciNum = mciNumArr[u];
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
                                            "reSanitize": false
                                        }));
                                        reXmlEq = {
                                            "reXmlEq": function (pixEq, jsStr, str) {
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
                                                    "jsStr": jsStr,
                                                    "str": str
                                                };
                                            }
                                        };
                                        reXmlEq.fileName = entry.fileName;
                                        str2 = ejs4xlx.parse(sheetBuf2, reXmlEq);
                                        anonymous = eval("(function anonymous(_args) {" + str2 + "})");
                                        buffer2 = anonymous.call(this, data);
                                        return _builder_$0.n(updateEntryAsync.apply(hzip, [entry.fileName, buffer2]), function () {
                                            (v = 0, len10 = imgTkArr.length)
                                            return _builder_$0.a(function() {
                                                return v < len10;
                                            }, function() {
                                                v ++;
                                            },
                                                _builder_$0.e(function() {
                                                    imgTk = imgTkArr[v];
                                                    return _builder_$0.n(imgTk, function () {
                                                        return _builder_$0.h();
                                                    });
                                                })
                                            );
                                        });
                                    });
                                })
                            );
                        }),
                        _builder_$0.e(function() {
                            sharedStrings2.push(new Buffer("</sst>"));
                            buffer2 = Buffer.concat(sharedStrings2);
                            sharedStrings2 = void 0;
                            return _builder_$0.n(updateEntryAsync.apply(hzip, ["xl/sharedStrings.xml", buffer2]), function () {
                                return _builder_$0.g(hzip.buffer);
                            });
                        })
                    );
                });
            });
        })
    );
});

  renderPath = (function (ejsDir, data) {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            var buffer, config, configPath, exists, exlBuf, extname, filter, ftObj, key, l, len1, obj, val;
            configPath = ejsDir + "/config.json";
            return _builder_$0.n(existsAsync(configPath), function (_result_$) {
                exists = _result_$;
                return _builder_$0.f(
                    _builder_$0.e(function() {
                        if (exists === false) {
                            extname = path.extname(ejsDir);
                            if (extname === "") {
                                ejsDir = ejsDir + ".xlsx";
                            }
                            return _builder_$0.n(readFileAsync(ejsDir), function (_result_$) {
                                exlBuf = _result_$;
                                return _builder_$0.n(renderExcel(exlBuf, data), function (_result_$) {
                                    return _builder_$0.g(_result_$);
                                });
                            });
                        } else {
                            return _builder_$0.h();
                        }
                    }),
                    _builder_$0.e(function() {
                        return _builder_$0.n(readFileAsync(configPath, "utf8"), function (_result_$) {
                            config = _result_$;
                            config = JSON.decode(config);
                            filter = [];
                            return _builder_$0.f(
                                _builder_$0.e(function() {
                                    (l = 0, len1 = config.length)
                                    return _builder_$0.a(function() {
                                        return l < len1;
                                    }, function() {
                                        l ++;
                                    },
                                        _builder_$0.e(function() {
                                            obj = config[l];
                                            key = obj.key;
                                            val = obj.value;
                                            if (key === void 0 || key === null) {
                                                return _builder_$0.j();
                                            }
                                            extname = path.extname(val);
                                            ftObj = {
                                                "path": val
                                            };
                                            if (extname !== ".xml" && extname !== ".rels") {
                                                ftObj.notEjs = true;
                                            }
                                            return _builder_$0.n(readFileAsync(ejsDir + "/" + key), function (_result_$) {
                                                ftObj.buffer = _result_$;
                                                filter.push(ftObj);
                                                return _builder_$0.h();
                                            });
                                        })
                                    );
                                }),
                                _builder_$0.e(function() {
                                    return _builder_$0.n(readFileAsync(ejsDir + "/" + path.basename(ejsDir) + ".zip"), function (_result_$) {
                                        buffer = _result_$;
                                        return _builder_$0.n(render(buffer, filter, data), function (_result_$) {
                                            return _builder_$0.g(_result_$);
                                        });
                                    });
                                })
                            );
                        });
                    })
                );
            });
        })
    );
});

  getExcelArrCb = function(buffer, callback) {
    var tmpFn;
    tmpFn = (function (buffer) {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            var rvObj;
            return _builder_$0.n(getExcelArr(buffer), function (_result_$) {
                rvObj = _result_$;
                if (callback) {
                    callback(rvObj);
                }
                return _builder_$0.h();
            });
        })
    );
});
    tmpFn(buffer).start();
  };

  getExcelArr = (function (buffer) {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            var buf, cEle, crStr, cs, enr, ens, entries, entry, fileName, hzip, i, l, len1, len2, len3, len4, m, n, numcr, numcrArr, o, p, ref2, ref3, row, sharedJson, sharedStr, sheet, sheetArr, sheetStr, sheets, sheetsEns, sir, vStr, vStr2;
            sharedStr = null;
            sheets = [];
            hzip = new Hzip(buffer);
            entries = hzip.entries;
            return _builder_$0.f(
                _builder_$0.e(function() {
                    (l = 0, len1 = entries.length)
                    return _builder_$0.a(function() {
                        return l < len1;
                    }, function() {
                        l ++;
                    },
                        _builder_$0.e(function() {
                            entry = entries[l];
                            fileName = entry.fileName;
                            if (fileName === "xl/sharedStrings.xml" || /xl\/worksheets\/sheet\d+\.xml/gm.test(fileName)) {
                                return _builder_$0.n(inflateRawAsync(entry.cfile), function (_result_$) {
                                    buf = _result_$;
                                    if (/xl\/worksheets\/sheet\d+\.xml/gm.test(fileName)) {
                                        sheets.push(buf);
                                    } else {
                                        sharedStr = buf;
                                    }
                                    return _builder_$0.h();
                                });
                            } else {
                                return _builder_$0.h();
                            }
                        })
                    );
                }),
                _builder_$0.e(function() {
                    sheetsEns = [];
                    sharedJson = xml2json.toJson(sharedStr);
                    sheetArr = [];
                    for ((m = 0, len2 = sheets.length); m < len2; m ++) {
                        sheetStr = sheets[m];
                        ens = [];
                        sheet = xml2json.toJson(sheetStr);
                        if (sheet.worksheet.sheetData.row === void 0) {
                            continue;
                        }
                        if (! isArray(sheet.worksheet.sheetData.row)) {
                            sheet.worksheet.sheetData.row = [sheet.worksheet.sheetData.row];
                        }
                        for ((i = n = 0, ref2 = sheet.worksheet.sheetData.row.length); (0 <= ref2) ? (n < ref2) : (n > ref2); i = (0 <= ref2) ? (++ n) : (-- n)) {
                            row = sheet.worksheet.sheetData.row[i];
                            if (! row.c) {
                                continue;
                            }
                            if (! isArray(row.c)) {
                                row.c = [row.c];
                            }
                            cs = row.c;
                            enr = [];
                            ens[parseInt(row.r) - 1] = enr;
                            numcrArr = [];
                            for ((o = 0, len3 = cs.length); o < len3; o ++) {
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
                                    if (! sharedJson.sst.si[vStr]) {
                                        continue;
                                    }
                                    if (sharedJson.sst.si[vStr].t) {
                                        vStr = sharedJson.sst.si[vStr].t["$t"];
                                    } else {
                                        vStr2 = "";
                                        if (! sharedJson.sst.si[vStr].r) {
                                            sharedJson.sst.si[vStr].r = [];
                                        } else if (! isArray(sharedJson.sst.si[vStr].r)) {
                                            sharedJson.sst.si[vStr].r = [sharedJson.sst.si[vStr].r];
                                        }
                                        ref3 = sharedJson.sst.si[vStr].r;
                                        for ((p = 0, len4 = ref3.length); p < len4; p ++) {
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
                    return _builder_$0.g(sheetArr);
                })
            );
        })
    );
});

  getExcelEns = function(sharedStr, sheets) {
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

  str2Xml = function(str) {
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

  charPlus = function(str, num) {
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

  charToNum = function(str) {
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

  exports.getExcelArrCb = getExcelArrCb;

}).call(this);
