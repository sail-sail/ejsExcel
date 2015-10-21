(function() {
  var Binding, Hzip, Task, Wind, charPlus, charToNum, crypto, drawingBuf2, drawingRelBuf2, ejs, ejs4xlx, err, existsAsync, fs, getExcelArr, getExcelArrCb, getExcelEns, inflateRawAsync, isArray, isFunction, isObject, isString, isType, path, readFileAsync, render, renderExcel, renderExcelCb, renderPath, replaceLast, sharedStrings2, sheetEntrieRel2, sheetSufStr, str2Xml, xjOp, xml2json, zlib;

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
      Wind = require("wind");
    } catch (_error) {
      err = _error;
      Wind = require("./lib/Wind");
    }
  }

  ejs4xlx = void 0;

  try {
    ejs4xlx = require("ejs4xlx");
  } catch (_error) {
    err = _error;
    ejs4xlx = require("./ejs4xlx");
  }

  ejs = void 0;

  try {
    ejs = require("ejs");
  } catch (_error) {
    err = _error;
    ejs = require("./lib/ejs");
  }

  Hzip = void 0;

  try {
    Hzip = require("hzip");
  } catch (_error) {
    err = _error;
    Hzip = require("./lib/hzip");
  }

  xml2json = void 0;

  try {
    xml2json = require("xml2json");
  } catch (_error) {
    err = _error;
    xml2json = require("./lib/xml2json");
  }

  Task = Wind.Async.Task;

  Binding = Wind.Async.Binding;

  existsAsync = Binding.fromCallback(fs.exists);

  readFileAsync = Binding.fromStandard(fs.readFile);

  inflateRawAsync = Binding.fromStandard(zlib.inflateRaw);

  render = (function (buffer, filter, _data_, hzip, options) {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    var _caller_$0 = this.$caller;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            this.$caller = _caller_$0;
            var anonymous, buffer2, data, entries, flt, l, len1, src, str, updateEntryAsync;
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
                "sharedStrings": []
            };
            data._ps_ = function (val) {
                var arr, index;
                val = str2Xml(val);
                arr = data._acVar_.sharedStrings;
                index = arr.indexOf(val);
                if (index === - 1) {
                    return arr.push(val) - 1;
                }
                return index;
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
                                        return _builder_$0.n(ejs.parse(flt.buffer), function (_result_$) {
                                            str = _result_$;
                                            anonymous = eval(Wind.compile("async", "function anonymous(_args) {\n" + str + "\n}"));
                                            return _builder_$0.n(anonymous.call(this, data), function (_result_$) {
                                                src = _result_$;
                                                buffer2 = new Buffer(src);
                                                return _builder_$0.h();
                                            });
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
                    return _builder_$0.g(hzip.buffer);
                })
            );
        })
    );
});

  sharedStrings2 = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\"><%\nvar _acVar_ = _args._acVar_;\nvar ssArr = _acVar_.sharedStrings;\nfor(var i=0; i<ssArr.length; i++) {\n$await(Wind.Async.sleep(0));\n%><si><t xml:space=\"preserve\"><%=ssArr[i]%></t><phoneticPr fontId=\"1\" type=\"noConversion\"/></si><%}%></sst>");

  sheetSufStr = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><%\nvar _data_ = _args._data_;\nvar _charPlus_ = _args._charPlus_;\nvar _charToNum_ = _args._charToNum_;\nvar _str2Xml_ = _args._str2Xml_;\nvar _ps_ = _args._ps_;\nvar _pi_ = _args._pi_;\nvar _pf_ = _args._pf_;\nvar _acVar_ = _args._acVar_;\nvar _r = 0;\nvar _c = 0;\nvar _row = 0;\nvar _col = \"\";\nvar _rc = \"\";\nvar _imgAsync_ = _args._imgAsync_;\nvar _img_ = _args._img_;\nvar _mergeCellArr_ = [];\n%>";

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
    var _caller_$0 = this.$caller;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            this.$caller = _caller_$0;
            var buf2;
            return _builder_$0.n(renderExcel(exlBuf, _data_), function (_result_$) {
                buf2 = _result_$;
                callback(buf2);
                return _builder_$0.h();
            });
        })
    );
});
    tmpFn().start();
  };

  renderExcel = (function (exlBuf, _data_) {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    var _caller_$0 = this.$caller;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            this.$caller = _caller_$0;
            var _imgFn_, anonymous, begin, buffer2, cItem, data, end, endElement, entry, hzip, i, imgTk, imgTkArr, l, len1, len2, len3, len4, len5, len6, len7, len8, m, m_c_i, mciNum, mciNumArr, mergeCell, n, o, p, phoneticPr, q, r, reXmlEq, ref, ref0, ref1, ref2, ref3, ref4, ref5, refArr, row, sharedStringsTmp2, sheetBuf, sheetBuf2, sheetDataElementState, sheetEntrieRels, sheetEntries, sheetObj, shsEntry, shsObj, shsStr, si, si2, sirTp, src2, startElement, str2, t, updateEntryAsync, xjOpTmp;
            data = {
                "_data_": _data_
            };
            data._charPlus_ = charPlus;
            data._charToNum_ = charToNum;
            data._str2Xml_ = str2Xml;
            data._acVar_ = {
                "sharedStrings": []
            };
            data._ps_ = function (str, buf) {
                var arr, i, index, l, ref, val;
                if (str === "") {
                    for (i = l = ref = buf.length - 1; (ref <= - 1) ? (l < - 1) : (l > - 1); i = (ref <= - 1) ? (++ l) : (-- l)) {
                        if (/<v>/gm.test(buf[i])) {
                            buf[i] = replaceLast(buf[i], /<v>/gm, "");
                            buf[i] = replaceLast(buf[i], /\s+t="s"/gm, "");
                            break;
                        }
                    }
                    buf.push = function (puhStr) {
                        var index;
                        index = - 1;
                        if (puhStr.indexOf("</v>") !== - 1) {
                            index = Array.prototype.push.apply(buf, [puhStr.replace("</v>", "")]);
                            buf.push = Array.prototype.push;
                        } else {
                            index = Array.prototype.push.apply(buf, [puhStr]);
                        }
                        return index;
                    };
                    return "";
                }
                val = str2Xml(str);
                arr = data._acVar_.sharedStrings;
                index = arr.indexOf(val);
                if (index === - 1) {
                    return arr.push(val) - 1;
                }
                return index;
            };
            data._pf_ = function (str, buf) {
                var i, l, m, ref, ref2;
                str = str2Xml(str);
                for (i = l = ref = buf.length - 1; (ref <= - 1) ? (l < - 1) : (l > - 1); i = (ref <= - 1) ? (++ l) : (-- l)) {
                    if (/<v>/gm.test(buf[i]) === true) {
                        buf[i] = replaceLast(buf[i], /<v>/gm, "<f>");
                        break;
                    }
                }
                for (i = m = ref2 = buf.length - 1; (ref2 <= - 1) ? (m < - 1) : (m > - 1); i = (ref2 <= - 1) ? (++ m) : (-- m)) {
                    if (/\s+t="s"/gm.test(buf[i]) === true) {
                        buf[i] = replaceLast(buf[i], /\s+t="s"/gm, "");
                        break;
                    }
                }
                buf.push = function (puhStr) {
                    if (puhStr.indexOf("</v>") !== - 1) {
                        Array.prototype.push.apply(buf, [puhStr.replace(/<\/v>/m, "</f>")]);
                        buf.push = Array.prototype.push;
                    } else {
                        Array.prototype.push.apply(buf, [puhStr]);
                    }
                    return this;
                };
                return str;
            };
            data._pi_ = function (str, buf) {
                var i, l, ref;
                if (isNaN(Number(str))) {
                    return data._ps_(str, buf);
                }
                str = str2Xml(str);
                for (i = l = ref = buf.length - 1; (ref <= - 1) ? (l < - 1) : (l > - 1); i = (ref <= - 1) ? (++ l) : (-- l)) {
                    if (/\s+t="s"/gm.test(buf[i]) === true) {
                        buf[i] = replaceLast(buf[i], /\s+t="s"/gm, "");
                        break;
                    }
                }
                return str;
            };
            data._acVar_ = {
                "sharedStrings": []
            };
            hzip = new Hzip(exlBuf);
            updateEntryAsync = Binding.fromStandard(hzip.updateEntry);
            return _builder_$0.n(updateEntryAsync.apply(hzip, ["xl/calcChain.xml"]), function () {
                sheetEntries = [];
                sheetEntrieRels = [];
                ref = hzip.entries;
                for ((l = 0, len1 = ref.length); l < len1; l ++) {
                    entry = ref[l];
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
                    var _caller_$0 = this.$caller;
                    return _builder_$1.m(this,
                        _builder_$1.e(function() {
                            this.$caller = _caller_$0;
                            var cfileName, drawingBuf, drawingObj, drawingRelBuf, drawingRelObj, drawingRelStr, drawingStr, entryImgTmp, entryTmp, eny, hashMd5, imgBaseName, imgBuf, imgPh, itHs, len2, len3, len4, len5, m, md5Str, n, o, p, ref2, ref3, ref4, ref5, sei, xdr_frt;
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
                                    ref2 = hzip.entries;
                                    for ((m = 0, len2 = ref2.length); m < len2; m ++) {
                                        entryTmp = ref2[m];
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
                                            ref3 = hzip.entries;
                                            return _builder_$1.f(
                                                _builder_$1.e(function() {
                                                    (n = 0, len3 = ref3.length)
                                                    return _builder_$1.a(function() {
                                                        return n < len3;
                                                    }, function() {
                                                        n ++;
                                                    },
                                                        _builder_$1.e(function() {
                                                            entryImgTmp = ref3[n];
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
                                                    ref4 = drawingRelObj["Relationships"]["Relationship"];
                                                    for ((o = 0, len4 = ref4.length); o < len4; o ++) {
                                                        eny = ref4[o];
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
                                                        ref5 = hzip.entries;
                                                        return _builder_$1.f(
                                                            _builder_$1.e(function() {
                                                                (p = 0, len5 = ref5.length)
                                                                return _builder_$1.a(function() {
                                                                    return p < len5;
                                                                }, function() {
                                                                    p ++;
                                                                },
                                                                    _builder_$1.e(function() {
                                                                        entryImgTmp = ref5[p];
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
                                        ref2 = sheetObj.worksheet.sheetData.row;
                                        for ((n = 0, len3 = ref2.length); n < len3; n ++) {
                                            row = ref2[n];
                                            if (row.c !== void 0) {
                                                if (! row.c) {
                                                    row.c = [];
                                                } else if (! isArray(row.c)) {
                                                    row.c = [row.c];
                                                }
                                                ref3 = row.c;
                                                for ((o = 0, len4 = ref3.length); o < len4; o ++) {
                                                    cItem = ref3[o];
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
                                                            ref4 = si.r;
                                                            for ((p = 0, len5 = ref4.length); p < len5; p ++) {
                                                                sirTp = ref4[p];
                                                                if (sirTp.t) {
                                                                    si2.t["$t"] += sirTp.t["$t"];
                                                                }
                                                            }
                                                        } else {
                                                            si2.t["$t"] = si.t["$t"];
                                                        }
                                                        cItem.v["$t"] = si2.t["$t"];
                                                        if (cItem.v !== void 0) {
                                                            if (! (cItem.v["$t"] === void 0 || cItem.v["$t"] === "")) {
                                                                begin = cItem.v["$t"].indexOf("<%");
                                                                end = cItem.v["$t"].indexOf("%>");
                                                                if (begin === - 1 || end === - 1) {
                                                                    cItem.v["$t"] = "<%='" + cItem.v["$t"].replace(/'/gm, "\\'") + "'%>";
                                                                }
                                                            }
                                                        }
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
                                                    if (sheetObj.worksheet.mergeCells !== void 0 && sheetObj.worksheet.mergeCells.mergeCell !== void 0) {
                                                        mciNumArr = [];
                                                        ref5 = sheetObj.worksheet.mergeCells.mergeCell;
                                                        for ((m_c_i = q = 0, len6 = ref5.length); q < len6; m_c_i = ++ q) {
                                                            mergeCell = ref5[m_c_i];
                                                            if (mergeCell.ref !== void 0) {
                                                                refArr = mergeCell.ref.split(":");
                                                                ref0 = refArr[0];
                                                                ref1 = refArr[1];
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
                                                }
                                            }
                                        }
                                        if (sheetObj.worksheet.mergeCells !== void 0) {
                                            sheetObj.worksheet.mergeCells = {
                                                "$t": "<% for(var m_cl=0; m_cl<_mergeCellArr_.length; m_cl++) { %><%-'<mergeCell ref=\"'+_mergeCellArr_[m_cl]+'\"/>'%><% } %>"
                                            };
                                        }
                                        sheetBuf2 = new Buffer(sheetSufStr + xml2json.toXml(sheetObj, "", {
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
                                                } else if (pixEq === "I") {
                                                    jsStr = jsStr.replace(/\n/gm, "\\n").replace(/\r/gm, "\\r").replace(/\t/gm, "\\t");
                                                    jsStr = "_img_(" + jsStr + ",buf)";
                                                }
                                                return {
                                                    "jsStr": jsStr,
                                                    "str": str
                                                };
                                            }
                                        };
                                        reXmlEq.fileName = entry.fileName;
                                        str2 = ejs4xlx.parse(sheetBuf2, reXmlEq);
                                        str2 = "(function (_args) {\n" + str2 + "\n})";
                                        anonymous = void 0;
                                        try {
                                            anonymous = eval(str2);
                                        } catch (_error) {
                                            err = _error;
                                            console.log(str2);
                                            return _builder_$0.k(err);
                                        }
                                        src2 = anonymous.call(this, data);
                                        buffer2 = new Buffer(src2);
                                        return _builder_$0.n(updateEntryAsync.apply(hzip, [entry.fileName, buffer2]), function () {
                                            (t = 0, len8 = imgTkArr.length)
                                            return _builder_$0.a(function() {
                                                return t < len8;
                                            }, function() {
                                                t ++;
                                            },
                                                _builder_$0.e(function() {
                                                    imgTk = imgTkArr[t];
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
                            sharedStringsTmp2 = ejs4xlx.parse(sharedStrings2);
                            sharedStringsTmp2 = "function anonymous(_args) {\n" + sharedStringsTmp2 + "\n}";
                            anonymous = eval(Wind.compile("async", sharedStringsTmp2));
                            return _builder_$0.n(anonymous.call(this, data), function (_result_$) {
                                sharedStringsTmp2 = _result_$;
                                buffer2 = new Buffer(sharedStringsTmp2);
                                return _builder_$0.n(updateEntryAsync.apply(hzip, ["xl/sharedStrings.xml", buffer2]), function () {
                                    return _builder_$0.g(hzip.buffer);
                                });
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
    var _caller_$0 = this.$caller;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            this.$caller = _caller_$0;
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
    var _caller_$0 = this.$caller;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            this.$caller = _caller_$0;
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
    var _caller_$0 = this.$caller;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            this.$caller = _caller_$0;
            var buf, cEle, crStr, cs, enr, ens, entries, entry, fileName, hzip, i, l, len1, len2, len3, len4, m, n, numcr, numcrArr, o, p, ref, ref2, row, sharedJson, sharedStr, sheet, sheetArr, sheetStr, sheets, sheetsEns, sir, vStr, vStr2;
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
                        for ((i = n = 0, ref = sheet.worksheet.sheetData.row.length); (0 <= ref) ? (n < ref) : (n > ref); i = (0 <= ref) ? (++ n) : (-- n)) {
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
                                if (cEle.t === "s") {
                                    if (sharedJson.sst.si[vStr].t !== void 0) {
                                        vStr = sharedJson.sst.si[vStr].t["$t"];
                                    } else {
                                        vStr2 = "";
                                        if (! sharedJson.sst.si[vStr].r) {
                                            sharedJson.sst.si[vStr].r = [];
                                        } else if (! isArray(sharedJson.sst.si[vStr].r)) {
                                            sharedJson.sst.si[vStr].r = [sharedJson.sst.si[vStr].r];
                                        }
                                        ref2 = sharedJson.sst.si[vStr].r;
                                        for ((p = 0, len4 = ref2.length); p < len4; p ++) {
                                            sir = ref2[p];
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
    var cEle, cont, crStr, cs, enr, ens, headsArr, i, k, l, len1, len2, len3, len4, m, n, numcr, numcrArr, o, p, q, ref, ref2, ref3, ref4, row, sharedJson, sheet, sheetHeadsArr, sheetStr, sheetsEns, sir, vStr, vStr2;
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
      ref = sheet.worksheet.sheetData.row[1].c;
      for (m = 0, len2 = ref.length; m < len2; m++) {
        cEle = ref[m];
        crStr = cEle.r.toString();
        crStr = crStr.replace("2", "");
        if (cEle.v === void 0) {
          continue;
        }
        vStr = cEle.v["$t"];
        if (cEle.t === "s") {
          vStr = sharedJson.sst.si[vStr].t["$t"];
        }
        numcr = charToNum(crStr);
        headsArr[numcr] = vStr;
      }
      for (i = n = 2, ref2 = sheet.worksheet.sheetData.row.length; 2 <= ref2 ? n < ref2 : n > ref2; i = 2 <= ref2 ? ++n : --n) {
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
          if (cEle.t === "s") {
            if (sharedJson.sst.si[vStr].t !== void 0) {
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
          if (vStr === "null" || vStr === "NULL") {
            vStr = null;
          }
          cont = headsArr[numcr];
          if (cont !== void 0) {
            enr[cont] = vStr;
          }
        }
        for (k = q = 0, ref4 = headsArr.length; 0 <= ref4 ? q < ref4 : q > ref4; k = 0 <= ref4 ? ++q : --q) {
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
    var charTmp, i, l, ref, s, str2;
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
    for (i = l = 0, ref = str.length; 0 <= ref ? l < ref : l > ref; i = 0 <= ref ? ++l : --l) {
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
    var i, j, l, len, ref, temp, val;
    str = new String(str);
    val = 0;
    len = str.length;
    for (j = l = 0, ref = len; 0 <= ref ? l < ref : l > ref; j = 0 <= ref ? ++l : --l) {
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
