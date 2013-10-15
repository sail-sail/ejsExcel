(function() {
  var Binding, Hzip, Task, Wind, charPlus, charToNum, ejs, ejs4xlx, existsAsync, fs, getExcelEns, inflateRawAsync, isArray, isFunction, isObject, isString, isType, path, readFileAsync, render, renderExcel, renderPath, replaceLast, sharedStrings2, sheetSufStr, str2Xml, xjOp, xml2json, zlib;

  require.extensions['.node_' + process.platform + "_" + process.arch] = function(module, filename) {
    return require.extensions['.node'](module, filename);
  };

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

  ejs4xlx = require("./ejs4xlx");

  ejs = require("./lib/ejs");

  Hzip = require("./lib/hzip");

  Wind = require("./lib/wind");

  xml2json = require("./lib/xml2json");

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
            var anonymous, buffer2, data, entries, flt, src, str, updateEntryAsync, _i, _len;
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
                    (_i = 0, _len = filter.length)
                    return _builder_$0.a(function() {
                        return _i < _len;
                    }, function() {
                        _i ++;
                    },
                        _builder_$0.e(function() {
                            flt = filter[_i];
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

  sharedStrings2 = new Buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">\n	<%\n	var _data_ = _args._data_;\n	var _charPlus_ = _args._charPlus_;\n	var _charToNum_ = _args._charToNum_;\n	var _str2Xml_ = _args._str2Xml_;\n	var _acVar_ = _args._acVar_;\n	var _ps_ = _args._ps_;\n	\n	//字符串\n	var ssArr = _acVar_.sharedStrings;\n	for(var i=0; i<ssArr.length; i++) {\n	$await(Wind.Async.sleep(0));\n	%>\n	<si>\n		<t xml:space=\"preserve\"><%=ssArr[i]%></t>\n		<phoneticPr fontId=\"1\" type=\"noConversion\"/>\n	</si>\n	<%}%>\n</sst>");

  sheetSufStr = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><%\nvar _data_ = _args._data_;\nvar _charPlus_ = _args._charPlus_;\nvar _charToNum_ = _args._charToNum_;\nvar _str2Xml_ = _args._str2Xml_;\nvar _ps_ = _args._ps_;\nvar _pi_ = _args._pi_;\nvar _pf_ = _args._pf_;\nvar _acVar_ = _args._acVar_;\nvar _r = 0;\nvar _c = 0;\n%>";

  xjOp = {
    object: true,
    reversible: true,
    coerce: false,
    trim: false,
    sanitize: false
  };

  exports.renderExcelCb = function(exlBuf, _data_, callback) {
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
            var anonymous, begin, buffer2, cItem, data, end, endElement, entries, entry, hzip, i, phoneticPr, reXmlEq, row, sharedStringsTmp2, sheetBuf, sheetBuf2, sheetDataElementState, sheetEntries, sheetObj, shsEntry, shsObj, shsStr, si, si2, sirTp, src2, startElement, str2, updateEntryAsync, xjOpTmp, _i, _j, _k, _l, _len, _len1, _len2, _len3, _len4, _m, _ref, _ref1, _ref2;
            data = {
                "_data_": _data_
            };
            data._charPlus_ = charPlus;
            data._charToNum_ = charToNum;
            data._str2Xml_ = str2Xml;
            data._acVar_ = {
                "sharedStrings": []
            };
            data._ps_ = function (str) {
                var arr, index, val;
                val = str2Xml(str);
                arr = data._acVar_.sharedStrings;
                index = arr.indexOf(val);
                if (index === - 1) {
                    return arr.push(val) - 1;
                }
                return index;
            };
            data._pf_ = function (str, buf) {
                var i, _i, _j, _ref, _ref1;
                str = str2Xml(str);
                for (i = _i = _ref = buf.length - 1; (_ref <= - 1) ? (_i < - 1) : (_i > - 1); i = (_ref <= - 1) ? (++ _i) : (-- _i)) {
                    if (/<v>/gm.test(buf[i]) === true) {
                        buf[i] = replaceLast(buf[i], /<v>/gm, "<f>");
                        break;
                    }
                }
                for (i = _j = _ref1 = buf.length - 1; (_ref1 <= - 1) ? (_j < - 1) : (_j > - 1); i = (_ref1 <= - 1) ? (++ _j) : (-- _j)) {
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
                var i, _i, _ref;
                for (i = _i = _ref = buf.length - 1; (_ref <= - 1) ? (_i < - 1) : (_i > - 1); i = (_ref <= - 1) ? (++ _i) : (-- _i)) {
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
            entries = hzip.entries;
            shsEntry = hzip.getEntry("xl/sharedStrings.xml");
            return _builder_$0.n(inflateRawAsync(shsEntry.cfile), function (_result_$) {
                shsStr = _result_$;
                shsObj = xml2json.toJson(shsStr);
                sheetEntries = [];
                for ((_i = 0, _len = entries.length); _i < _len; _i ++) {
                    entry = entries[_i];
                    if (entry.fileName.indexOf("xl/worksheets/sheet") === 0) {
                        sheetEntries.push(entry);
                    }
                }
                sheetEntries.sort(function (arg0, arg1) {
                    return arg0.fileName > arg1.fileName;
                });
                return _builder_$0.f(
                    _builder_$0.e(function() {
                        (i = _j = 0, _len1 = sheetEntries.length)
                        return _builder_$0.a(function() {
                            return _j < _len1;
                        }, function() {
                            i = ++ _j;
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
                                        if (isString(value)) {
                                            return value;
                                        }
                                        if (sheetDataElementState === "start") {
                                            return value;
                                        }
                                        value = value.replace(/[&<>]/gm, function (g1, g2) {
                                            if (g1 === "&") {
                                                return "&amp;";
                                            } else if (g1 === "<") {
                                                return "&lt;";
                                            } else if (g1 === ">") {
                                                return "&gt;";
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
                                    _ref = sheetObj.worksheet.sheetData.row;
                                    for ((_k = 0, _len2 = _ref.length); _k < _len2; _k ++) {
                                        row = _ref[_k];
                                        if (row.c === void 0) {
                                            continue;
                                        } else if (! isArray(row.c)) {
                                            row.c = [row.c];
                                        }
                                        _ref1 = row.c;
                                        for ((_l = 0, _len3 = _ref1.length); _l < _len3; _l ++) {
                                            cItem = _ref1[_l];
                                            if (cItem.t === "s" && cItem.v !== void 0) {
                                                if (! isArray(shsObj.sst.si)) {
                                                    shsObj.sst.si = [shsObj.sst.si];
                                                }
                                                si = shsObj.sst.si[cItem.v["$t"]];
                                                phoneticPr = si.phoneticPr;
                                                si2 = {
                                                    "t": {
                                                        "$t": ""
                                                    }
                                                };
                                                if (isArray(si.r)) {
                                                    _ref2 = si.r;
                                                    for ((_m = 0, _len4 = _ref2.length); _m < _len4; _m ++) {
                                                        sirTp = _ref2[_m];
                                                        si2.t["$t"] += sirTp.t["$t"];
                                                    }
                                                } else {
                                                    si2.t["$t"] = si.t["$t"];
                                                }
                                                cItem.v["$t"] = si2.t["$t"];
                                                if (cItem.v !== void 0) {
                                                    if (cItem.v["$t"] === void 0 || cItem.v["$t"] === "") {
                                                        continue;
                                                    }
                                                    begin = cItem.v["$t"].indexOf("<%");
                                                    end = cItem.v["$t"].indexOf("%>");
                                                    if (begin === - 1 || end === - 1) {
                                                        cItem.v["$t"] = "<%='" + cItem.v["$t"].replace(/'/gm, "\\'") + "'%>";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    sheetBuf2 = new Buffer(sheetSufStr + xml2json.toXml(sheetObj, "", {
                                        "reSanitize": false
                                    }));
                                    reXmlEq = {
                                        "reXmlEq": function (pixEq, jsStr, str) {
                                            if (pixEq === "=") {
                                                jsStr = "_ps_(" + jsStr + ")";
                                            } else if (pixEq === "~") {
                                                jsStr = "_pi_(" + jsStr + ",buf)";
                                            } else if (pixEq === "#") {
                                                jsStr = "_pf_(" + jsStr + ",buf)";
                                            }
                                            return {
                                                "jsStr": jsStr,
                                                "str": str
                                            };
                                        }
                                    };
                                    str2 = ejs4xlx.parse(sheetBuf2, reXmlEq);
                                    str2 = "function anonymous(_args) {\n" + str2 + "\n}";
                                    anonymous = eval(Wind.compile("async", str2));
                                    return _builder_$0.n(anonymous.call(this, data), function (_result_$) {
                                        src2 = _result_$;
                                        buffer2 = new Buffer(src2);
                                        return _builder_$0.n(updateEntryAsync.apply(hzip, [entry.fileName, buffer2]), function () {
                                            return _builder_$0.h();
                                        });
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
            var buffer, config, configPath, exists, exlBuf, extname, filter, ftObj, key, obj, val, _i, _len;
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
                                    (_i = 0, _len = config.length)
                                    return _builder_$0.a(function() {
                                        return _i < _len;
                                    }, function() {
                                        _i ++;
                                    },
                                        _builder_$0.e(function() {
                                            obj = config[_i];
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

  getExcelEns = function(sharedStr, sheets) {
    var cEle, cont, crStr, cs, enr, ens, headsArr, i, k, numcr, numcrArr, row, sharedJson, sheet, sheetStr, sheetsEns, sir, vStr, vStr2, _i, _j, _k, _l, _len, _len1, _len2, _len3, _m, _n, _ref, _ref1, _ref2, _ref3;
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
    for (_i = 0, _len = sheets.length; _i < _len; _i++) {
      sheetStr = sheets[_i];
      ens = [];
      if (isString(sheetStr) || Buffer.isBuffer(sheetStr)) {
        sheet = xml2json.toJson(sheetStr);
      } else {
        sheet = sheetStr;
      }
      row = sheet.worksheet.sheetData.row;
      headsArr = [];
      _ref = sheet.worksheet.sheetData.row[1].c;
      for (_j = 0, _len1 = _ref.length; _j < _len1; _j++) {
        cEle = _ref[_j];
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
      for (i = _k = 2, _ref1 = sheet.worksheet.sheetData.row.length; 2 <= _ref1 ? _k < _ref1 : _k > _ref1; i = 2 <= _ref1 ? ++_k : --_k) {
        row = sheet.worksheet.sheetData.row[i];
        cs = row.c;
        enr = {};
        ens.push(enr);
        numcrArr = [];
        for (_l = 0, _len2 = cs.length; _l < _len2; _l++) {
          cEle = cs[_l];
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
              if (!isArray(sharedJson.sst.si[vStr].r)) {
                sharedJson.sst.si[vStr].r = [sharedJson.sst.si[vStr].r];
              }
              _ref2 = sharedJson.sst.si[vStr].r;
              for (_m = 0, _len3 = _ref2.length; _m < _len3; _m++) {
                sir = _ref2[_m];
                if (sir.t === void 0) {
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
        for (k = _n = 0, _ref3 = headsArr.length; 0 <= _ref3 ? _n < _ref3 : _n > _ref3; k = 0 <= _ref3 ? ++_n : --_n) {
          if (numcrArr.indexOf(k) !== -1) {
            continue;
          }
          enr[headsArr[k]] = "";
        }
      }
      sheetsEns.push(ens);
    }
    return sheetsEns;
  };

  str2Xml = function(str) {
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

    str = str.replace(/[&<>]/gm, function(s) {
      if (s === "&") {
        return "&amp;";
      } else if (s === "<") {
        return "&lt;";
      } else if (s === ">") {
        return "&gt;";
      }
      return s;
    });
    return str;
  };

  charPlus = function(str, num) {
    var arr, i, i_$_, k, minus, tmp, _i, _j, _ref;
    if (num === 0) {
      return str;
    }
    arr = [];
    for (i_$_ = _i = 0, _ref = str.length; 0 <= _ref ? _i < _ref : _i > _ref; i_$_ = 0 <= _ref ? ++_i : --_i) {
      arr[i_$_] = str.charAt(i_$_);
    }
    minus = false;
    if (num < 0) {
      minus = true;
      num = Math.abs(num);
    }
    for (k = _j = 0; 0 <= num ? _j < num : _j > num; k = 0 <= num ? ++_j : --_j) {
      i = 0;
      while (true) {
        tmp = "";
        if (minus) {
          tmp = "A";
        } else {
          tmp = "Z";
        }
        if (arr[arr.length - 1 - i] !== tmp) {
          if (arr.length - 1 - i >= 0) {
            tmp = typeof minus === "function" ? minus(-{
              1: 1
            }) : void 0;
            if (minus) {
              tmp = -1;
            } else {
              tmp = 1;
            }
            arr[arr.length - 1 - i] = String.fromCharCode(arr[arr.length - 1 - i].charCodeAt(0) + tmp);
          } else {
            if (minus) {
              arr.shift();
            } else {
              arr.unshift("A");
            }
          }
          break;
        } else {
          if (minus) {
            arr[arr.length - 1 - i] = "Z";
            i++;
          } else {
            arr[arr.length - 1 - i] = "A";
            i++;
          }
          continue;
        }
      }
    }
    return arr.join("");
  };

  charToNum = function(str) {
    var code, i, num, _i, _ref;
    str = new String(str);
    num = 0;
    for (i = _i = 0, _ref = str.length; 0 <= _ref ? _i < _ref : _i > _ref; i = 0 <= _ref ? ++_i : --_i) {
      code = str.charCodeAt(i);
      num += code - 65 + (str.length - 1 - i) * 26;
    }
    return num;
  };

  exports.charPlus = charPlus;

  exports.charToNum = charToNum;

  exports.renderPath = renderPath;

  exports.render = render;

  exports.getExcelEns = getExcelEns;

  exports.renderExcel = renderExcel;

}).call(this);
