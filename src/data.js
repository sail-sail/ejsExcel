const fs = require("fs");
const path = require("path");
const Stream = require("stream");
const zlib = require("zlib");
const crypto = require("crypto");

// third-pard dependcies
const Hzip = require("../lib/hzip");
const xml2json = require("../lib/xml2json");
const xmldom = require("../lib/xmldom");
const qr = require("../lib/qr-image");


// subpackage dependencies
const ejs4xlx = require("./ejs4xlx");
const {isType,isObject,isArray,isFunction,isString}=require('./is-type');
const {charToNum,charPlus,str2Xml,replaceLast,stripWhitespace}=require('./utils');
const {Promise_fromCallback,Promise_fromStandard,Promise_sleep}=require('./async.js');


const DOMParser = xmldom.DOMParser;
const sheetSufStr = Buffer.from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<%\nvar _data_ = _args._data_;\nvar _charPlus_ = _args._charPlus_;\nvar _charToNum_ = _args._charToNum_;\nvar _str2Xml_ = _args._str2Xml_;\nvar _hideSheet_ = _args._hideSheet_;\nvar _showSheet_ = _args._showSheet_;\nvar _deleteSheet_ = _args._deleteSheet_;\nvar _ps_ = _args._ps_;\nvar _pi_ = _args._pi_;\nvar _pf_ = _args._pf_;\nvar _acVar_ = _args._acVar_;\nvar _r = 0;\nvar _c = 0;\nvar _row = 0;\nvar _col = \"\";\nvar _rc = \"\";\nvar _img_ = _args._img_;\nvar _qrcode_ = _args._qrcode_;\nvar _mergeCellArr_ = [];\nvar _mergeCellFn_ = function(mclStr) {\n	_mergeCellArr_.push(mclStr);\n};\nvar _hyperlinkArr_ = [];\nvar _outlineLevel_ = _args._outlineLevel_;\n%>");

const sharedStrings2Prx = Buffer.from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">");

var xjOp = {
    object: true,
    reversible: true,
    coerce: false,
    trim: false,
    sanitize: false
};


var existsAsync = Promise_fromCallback(fs.exists, fs);

var readFileAsync = Promise_fromStandard(fs.readFile, fs);

var inflateRawAsync = Promise_fromStandard(zlib.inflateRaw, zlib);

function DataFactory(exlBuf,_data_){

    const d= {
        _data_,
        _acVar_ : {
            sharedStrings: [],
            _ss_len: 0,
        },
        sharedStrings2 : [sharedStrings2Prx],
        hzip : new Hzip(exlBuf),
        _charPlus_:charPlus,
        _charToNum_:charToNum,
        _str2Xml_:str2Xml,
    }
    return d;

}

/**
 * 把元素数组`elementArray`中的所有元素的`r`属性交换到第一个属性位置
 * @param {Array<RowElement>} elementArray 
 */
function normalizeRAttributePosition(elementArray){
    // 迭代数组中的所有元素，把元素的"r"属性交换到第一个属性位置
    for (let n = 0; n < elementArray; n++) {
        let element = elementArray[n];
        if (element.attributes[0] && element.attributes[0].nodeName === "r") {
            continue;
        }
        // 迭代该元素的所有属性，找到r属性位置索引
        let attributes = element.attributes;
        // r属性
        let attr_r = null;
        // r属性所在的位置索引
        let attr_r_idx = null;
        for (let i=0; i <attributes.length ; i++ ) {
            let attr = attributes[i];
            if (attr.nodeName === "r") {
                attr_r = attr;
                attr_r_idx = i;
                break;
            }
        }
        if (!attr_r) { continue; }

        // 交换 row.r 属性到第一个位置
        let attr0 = element.attributes[0];
        element.attributes[0] = attr_r;
        element.attributes[attr_r_idx] = attr0;
    }
}


const renderExcel = async function (exlBuf, _data_) {
    var anonymous, attr, attr0, attr_r, begin, buffer2, cEl, cElArr, cItem,  end, endElement, entry, hyperlink, hyperlinksDomEl, i, i1, idx, j1, key, keyArr,  len10, len11, len12, len13, len14, len15, len3, len4, len5, len6, len7, len8, len9, m, m_c_i, mciNum, mciNumArr, mergeCell, n, o, p, pageMarginsDomEl, phoneticPr, phoneticPrDomEl, q, r, reXmlEq, ref, ref0, ref1, ref2, ref3, ref4, ref5, ref6, ref7, ref8, ref9, refArr, row, rowEl, rowElArr, sheetBuf, sheetBuf2, sheetDataElementState,  sheetObj,  si, si2, sirTp, startElement, str, str2, u,  v, w,  x, xjOpTmp, y, z;
    let data = DataFactory(exlBuf,_data_);
    let sharedStrings2 = [sharedStrings2Prx];
    let hzip =data.hzip;
    let updateEntryAsync = Promise_fromStandard(hzip.updateEntry, hzip);
    await updateEntryAsync("xl/calcChain.xml");
    ref2 = hzip.entries;
    let sheetEntries = ref2.filter(e=> e.fileName.indexOf("xl/worksheets/sheet") === 0 )
        .sort((arg0, arg1)=> arg0.fileName > arg1.fileName);
    let sheetEntrieRels = ref2.filter(e=> e.fileName.indexOf("xl/worksheets/_rels/") === 0)
        .sort((arg0, arg1) =>arg0.fileName > arg1.fileName);
    let workbookEntry = hzip.getEntry("xl/workbook.xml");
    let workbookBuf = (await inflateRawAsync(workbookEntry.cfile));
    let workbookRelsEntry = hzip.getEntry("xl/_rels/workbook.xml.rels");
    let workbookRelsBuf = (await inflateRawAsync(workbookRelsEntry.cfile));

    data._ps_ = function (str, buf) {
        var i, index, l, ref2, tmpStr, val;
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
                    buf[i] = Buffer.from(replaceLast(replaceLast(tmpStr, /<v>/gm, ""), /\s+t="s"/gm, ""));
                    break;
                }
            }
            buf.push = function (puhStr) {
                var index;
                puhStr = puhStr.toString();
                index = -1;
                if (puhStr.indexOf("</v>") !== -1) {
                    index = Array.prototype.push.apply(buf, [Buffer.from(puhStr.replace("</v>", ""))]);
                    buf.push = Array.prototype.push;
                } else {
                    index = Array.prototype.push.apply(buf, [Buffer.from(puhStr)]);
                }
                return index;
            };
            return "";
        }
        val = str2Xml(str);
        sharedStrings2.push(Buffer.from("<si><t xml:space=\"preserve\">" + val + "</t></si>"));
        index = data._acVar_._ss_len;
        data._acVar_._ss_len++;
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
                buf[i] = Buffer.from(replaceLast(tmpStr, /<v>/gm, "<f>"));
                break;
            }
        }
        for (i = m = ref3 = buf.length - 1; ref3 <= -1 ? m < -1 : m > -1; i = ref3 <= -1 ? ++m : --m) {
            tmpStr = buf[i].toString();
            if (/\s+t="s"/gm.test(tmpStr) === true) {
                buf[i] = Buffer.from(replaceLast(tmpStr, /\s+t="s"/gm, ""));
                break;
            }
        }
        buf.push = function (puhStr) {
            var index;
            puhStr = puhStr.toString();
            if (puhStr.indexOf("</v>") !== -1) {
                index = Array.prototype.push.apply(buf, [Buffer.from(puhStr.replace(/<\/v>/m, "</f>"))]);
                buf.push = Array.prototype.push;
            } else {
                index = Array.prototype.push.apply(buf, [Buffer.from(puhStr)]);
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
                buf[i] = Buffer.from(replaceLast(tmpStr, /\s+t="s"/gm, ""));
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
                buf[i] = Buffer.from(replaceLast(tmpStr, /<row\s/gm, "<row outlineLevel=\"" + str + "\" "));
                break;
            }
        }
    };
    data._img_ = async function (imgOpt, fileName, rowNum, cellNum) {
        var cNvPrDescr, cNvPrName, cfileName, doc, documentElement, drawingBuf, drawingEl, drawingObj, drawingRId, drawingRelBuf, drawingRelObj, drawingRelStr, drawingStr, entryImgTmp, entryTmp, eny, err, hashMd5, imgBaseName, imgBuf, imgPh, itHs, len2, len3, len4, len5, len6, len7, len8, m, md5Str, n, o, p, q, r, ref3, ref4, ref5, ref6, relationshipEl, relationshipElArr, sei, sheetEntrieRel, sheetEntry, sheetRelPth, shipEl, u, xdr_frt;
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
        if (isString(imgPh)) {
            try {
                imgBuf = (await readFileAsync(imgPh));
            } catch (error) {
                err = error;
                return "";
            }
            imgBaseName = path.basename(imgPh);
        } else if (Buffer.isBuffer(imgPh)) {
            imgBuf = imgPh;
        } else if (imgPh instanceof Stream) {
            imgBuf = (await new Promise(function (resolve, reject) {
                var imgBufArr;
                imgBufArr = [];
                imgPh.on("data", function (bufTmp) {
                    imgBufArr.push(bufTmp);
                });
                imgPh.on("end", function () {
                    resolve(Buffer.concat(imgBufArr));
                });
                imgPh.on("error", function (err) {
                    reject(err);
                });
            }));
        } else {
            return "";
        }
        hashMd5 = crypto.createHash("md5");
        md5Str = hashMd5.update(imgBuf).digest("hex");
        md5Str = "a" + md5Str;
        if (!imgBaseName) {
            imgBaseName = md5Str + ".png";
        }
        cNvPrName = imgOpt.cNvPrName || imgBaseName;
        cNvPrDescr = imgOpt.cNvPrDescr || imgBaseName;
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
            await updateEntryAsync.apply(hzip, [cfileName, imgBuf, false]);
        }
        entryTmp = void 0;
        for (n = 0, len3 = sheetEntries.length; n < len3; n++) {
            sheetEntry = sheetEntries[n];
            if (fileName !== sheetEntry.fileName) {
                continue;
            }
            entryTmp = sheetEntry;
            break;
        }
        if (!entry) {
            throw new Error(fileName + " dose not found!");
        }
        doc = new DOMParser().parseFromString(((await inflateRawAsync(entryTmp.cfile))).toString(), 'text/xml');
        documentElement = doc.documentElement;
        drawingEl = documentElement.getElementsByTagName("drawing")[0];
        if (!drawingEl) {
            throw "Excel模板显示动态图片之前,至少需要插入一张1像素的透明的图片,以初始化";
        }
        drawingRId = drawingEl.getAttribute("r:id");
        sheetEntrieRel = void 0;
        sheetRelPth = path.dirname(fileName) + "/_rels/" + path.basename(fileName) + ".rels";
        for (o = 0, len4 = sheetEntrieRels.length; o < len4; o++) {
            entryTmp = sheetEntrieRels[o];
            if (entryTmp.fileName === sheetRelPth) {
                sheetEntrieRel = entryTmp;
                break;
            }
        }
        if (!sheetEntrieRel) {
            throw new Error(sheetRelPth + " dose not found!");
        }
        doc = new DOMParser().parseFromString(((await inflateRawAsync(sheetEntrieRel.cfile))).toString(), 'text/xml');
        documentElement = doc.documentElement;
        relationshipElArr = documentElement.getElementsByTagName("Relationship");
        relationshipEl = void 0;
        for (p = 0, len5 = relationshipElArr.length; p < len5; p++) {
            shipEl = relationshipElArr[p];
            if (drawingRId === shipEl.getAttribute("Id")) {
                relationshipEl = shipEl;
                break;
            }
        }
        if (!relationshipEl) {
            throw new Error(sheetRelPth + (" Relationship.Id " + drawingRId + " dose not found!"));
        }
        sei = path.basename(relationshipEl.getAttribute("Target")).replace("drawing", "").replace(".xml", "");
        sei = Number(sei) - 1;
        drawingRelBuf = void 0;
        ref4 = hzip.entries;
        for (q = 0, len6 = ref4.length; q < len6; q++) {
            entryImgTmp = ref4[q];
            if (entryImgTmp.fileName === "xl/drawings/_rels/drawing" + (sei + 1) + ".xml.rels") {
                drawingRelBuf = (await inflateRawAsync(entryImgTmp.cfile));
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
        for (r = 0, len7 = ref5.length; r < len7; r++) {
            eny = ref5[r];
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
        await updateEntryAsync.apply(hzip, ["xl/drawings/_rels/drawing" + (sei + 1) + ".xml.rels", Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + drawingRelStr)]);
        drawingBuf = void 0;
        ref6 = hzip.entries;
        for (u = 0, len8 = ref6.length; u < len8; u++) {
            entryImgTmp = ref6[u];
            if (entryImgTmp.fileName === "xl/drawings/drawing" + (sei + 1) + ".xml") {
                drawingBuf = (await inflateRawAsync(entryImgTmp.cfile));
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
                        "name": cNvPrName,
                        "descr": cNvPrDescr
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
        await updateEntryAsync.apply(hzip, ["xl/drawings/drawing" + (sei + 1) + ".xml", Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + drawingStr)]);
        return "";
    };
    data._qrcode_ = async function (imgOpt, fileName, rowNum, cellNum) {
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
        await new Promise(function (resolve, reject) {
            imgSt.on("error", function (err) {
                return reject(err);
            });
            imgSt.on("end", function () {
                return resolve();
            });
        });
        imgOpt.imgPh = Buffer.concat(qrBufArr);
        rvObj = (await data._img_(imgOpt, fileName, rowNum, cellNum));
        return rvObj;
    };

    data._hideSheet_ = function (fileName) {
        var doc, documentElement, len2, len3, m, n, rId, relationshipEl, relationshipElArr, sheetEl, sheetElArr, sheetsEl;
        if (!workbookEntry || !workbookRelsEntry || !fileName) {
            return;
        }
        rId = void 0;
        doc = new DOMParser().parseFromString(workbookRelsBuf.toString(), 'text/xml');
        documentElement = doc.documentElement;
        relationshipElArr = documentElement.getElementsByTagName("Relationship");
        for (m = 0, len2 = relationshipElArr.length; m < len2; m++) {
            relationshipEl = relationshipElArr[m];
            if ("xl/" + relationshipEl.getAttribute("Target") === fileName) {
                rId = relationshipEl.getAttribute("Id");
                break;
            }
        }
        if (rId) {
            doc = new DOMParser().parseFromString(workbookBuf.toString(), 'text/xml');
            documentElement = doc.documentElement;
            sheetsEl = documentElement.getElementsByTagName("sheets")[0];
            sheetElArr = sheetsEl.getElementsByTagName("sheet");
            for (n = 0, len3 = sheetElArr.length; n < len3; n++) {
                sheetEl = sheetElArr[n];
                if (sheetEl.getAttribute("r:id") === rId) {
                    sheetEl.setAttribute("state", "hidden");
                    break;
                }
            }
            workbookBuf = Buffer.from(doc.toString());
        }
    };
    data._showSheet_ = function (fileName) {
        var doc, documentElement, len2, len3, m, n, rId, relationshipEl, relationshipElArr, sheetEl, sheetElArr, sheetsEl;
        if (!workbookEntry || !workbookRelsEntry || !fileName) {
            return;
        }
        rId = void 0;
        doc = new DOMParser().parseFromString(workbookRelsBuf.toString(), 'text/xml');
        documentElement = doc.documentElement;
        relationshipElArr = documentElement.getElementsByTagName("Relationship");
        for (m = 0, len2 = relationshipElArr.length; m < len2; m++) {
            relationshipEl = relationshipElArr[m];
            if ("xl/" + relationshipEl.getAttribute("Target") === fileName) {
                rId = relationshipEl.getAttribute("Id");
                break;
            }
        }
        if (rId) {
            doc = new DOMParser().parseFromString(workbookBuf.toString(), 'text/xml');
            documentElement = doc.documentElement;
            sheetsEl = documentElement.getElementsByTagName("sheets")[0];
            sheetElArr = sheetsEl.getElementsByTagName("sheet");
            for (n = 0, len3 = sheetElArr.length; n < len3; n++) {
                sheetEl = sheetElArr[n];
                if (sheetEl.getAttribute("r:id") === rId) {
                    sheetEl.removeAttribute("state");
                    break;
                }
            }
            workbookBuf = Buffer.from(doc.toString());
        }
    };
    data._deleteSheet_ = function (fileName) {
        var activeTab, bookViewsEl, delRelationshipEl, delSheet, doc, documentElement, len2, len3, len4, m, n, o, rId, relationshipEl, relationshipElArr, sheetEl, sheetElArr, sheetEntry, sheetsEl, workbookViewEl;
        if (!workbookEntry || !workbookRelsEntry || !fileName) {
            return;
        }
        rId = void 0;
        delRelationshipEl = void 0;
        doc = new DOMParser().parseFromString(workbookRelsBuf.toString(), 'text/xml');
        documentElement = doc.documentElement;
        relationshipElArr = documentElement.getElementsByTagName("Relationship");
        for (m = 0, len2 = relationshipElArr.length; m < len2; m++) {
            relationshipEl = relationshipElArr[m];
            if ("xl/" + relationshipEl.getAttribute("Target") === fileName) {
                delRelationshipEl = relationshipEl;
                rId = relationshipEl.getAttribute("Id");
                break;
            }
        }
        if (delRelationshipEl) {
            documentElement.removeChild(delRelationshipEl);
        }
        workbookRelsEntry = Buffer.from(doc.toString());
        doc = new DOMParser().parseFromString(workbookBuf.toString(), 'text/xml');
        documentElement = doc.documentElement;
        sheetsEl = documentElement.getElementsByTagName("sheets")[0];
        sheetElArr = sheetsEl.getElementsByTagName("sheet");
        delSheet = void 0;
        for (n = 0, len3 = sheetElArr.length; n < len3; n++) {
            sheetEl = sheetElArr[n];
            if (sheetEl.getAttribute("r:id") === rId) {
                delSheet = sheetEl;
                break;
            }
        }
        if (delSheet) {
            sheetsEl.removeChild(delSheet);
            sheetElArr = sheetsEl.getElementsByTagName("sheet");
            bookViewsEl = documentElement.getElementsByTagName("bookViews")[0];
            if (bookViewsEl) {
                workbookViewEl = bookViewsEl.getElementsByTagName("workbookView")[0];
                if (workbookViewEl) {
                    activeTab = workbookViewEl.getAttribute("activeTab");
                    if (activeTab) {
                        activeTab = Number(activeTab);
                        if (activeTab > sheetElArr.length) {
                            workbookViewEl.setAttribute("activeTab", sheetElArr.length);
                        }
                    }
                }
            }
        }
        workbookBuf = Buffer.from(doc.toString());
        for (o = 0, len4 = sheetEntries.length; o < len4; o++) {
            sheetEntry = sheetEntries[o];
            if (sheetEntry.fileName === fileName) {
                sheetEntry.__remove_sheet = true;
                break;
            }
        }
    };
    let shsEntry = hzip.getEntry("xl/sharedStrings.xml");
    if (shsEntry === void 0) { return exlBuf; }

    let shsStr = (await inflateRawAsync(shsEntry.cfile));
    let shsObj = xml2json.toJson(shsStr);
    // 迭代所有sheets
    for (i = m = 0 ; m < sheetEntries.length; i = ++m) {
        entry = sheetEntries[i];
        sheetBuf = (await inflateRawAsync(entry.cfile));

        let doc = new DOMParser().parseFromString(sheetBuf.toString(), 'text/xml');
        let documentElement = doc.documentElement;

        // 检测sheetData部分是否为空
        let sheetDataDomEl = documentElement.getElementsByTagName("sheetData")[0];
        if (!sheetDataDomEl) { continue; }

        // 检测是否有合并单元格
        let mergeCellsDomEl = documentElement.getElementsByTagName("mergeCells")[0];
        if (!mergeCellsDomEl) {
            mergeCellsDomEl = doc.createElement("mergeCells");
            hyperlinksDomEl = documentElement.getElementsByTagName("hyperlinks")[0];
            phoneticPrDomEl = documentElement.getElementsByTagName("phoneticPr")[0];
            pageMarginsDomEl = documentElement.getElementsByTagName("pageMargins")[0];
            if (phoneticPrDomEl) {
                documentElement.insertBefore(mergeCellsDomEl, phoneticPrDomEl);
            } else if (hyperlinksDomEl) {
                documentElement.insertBefore(mergeCellsDomEl, hyperlinksDomEl);
            } else if (pageMarginsDomEl) {
                documentElement.insertBefore(mergeCellsDomEl, pageMarginsDomEl);
            } else {
                documentElement.insertBefore(mergeCellsDomEl, sheetDataDomEl.nextSibling);
            }
        }

        // 迭代所有行，把<row/>元素的"r"属性交换到第一个属性位置
        rowElArr = sheetDataDomEl.getElementsByTagName("row");
        normalizeRAttributePosition(rowElArr);
        // 迭代所有单元格，把<c/>元素的"r"属性交换到第一个属性位置
        cElArr = sheetDataDomEl.getElementsByTagName("c");
        normalizeRAttributePosition(cElArr);

        sheetBuf = Buffer.from(doc.toString());
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
                            if (charToNum(cItem.r.replace(/\d+/, "")) >= charToNum(ref.replace(/\d+/, "")) && Number(cItem.r.replace(/\D+/, "")) >= Number(ref.replace(/\D+/, ""))) {
                                if (cItem.v != null) {
                                    if (!cItem.v["$t"]) {
                                        cItem.v["$t"] = "";
                                    }
                                    cItem.v["$t"] += "<% _hyperlinkArr_.push({ref:_charPlus_('" + ref.replace(/\d+/, '') + "',_c)+(" + Number(ref.replace(/\D+/, "")) + "+_r)";
                                    keyArr = Object.keys(hyperlink);
                                    for (z = 0, len13 = keyArr.length; z < len13; z++) {
                                        key = keyArr[z];
                                        if (key === "ref" || key === "display") {
                                            continue;
                                        }
                                        cItem.v["$t"] += (",'" + key + "':'") + hyperlink[key].replace(/'/gm, "\\'").replace(/\n/gm, "\\n") + "'";
                                    }
                                    cItem.v["$t"] += "})%>";
                                } else {
                                    if (!cItem["$t"]) {
                                        cItem["$t"] = "";
                                    }
                                    cItem["$t"] += "<% _hyperlinkArr_.push({ref:_charPlus_('" + ref.replace(/\d+/, '') + "',_c)+(" + Number(ref.replace(/\D+/, "")) + "+_r)";
                                    keyArr = Object.keys(hyperlink);
                                    for (i1 = 0, len14 = keyArr.length; i1 < len14; i1++) {
                                        key = keyArr[i1];
                                        if (key === "ref" || key === "display") {
                                            continue;
                                        }
                                        cItem["$t"] += (",'" + key + "':'") + hyperlink[key].replace(/'/gm, "\\'").replace(/\n/gm, "\\n") + "'";
                                    }
                                    cItem["$t"] += "})%>";
                                }
                                mciNumArr.push(m_c_i);
                            }
                        }
                        for (j1 = 0, len15 = mciNumArr.length; j1 < len15; j1++) {
                            mciNum = mciNumArr[j1];
                            sheetObj.worksheet.hyperlinks.hyperlink.splice(mciNum, 1);
                        }
                    }
                }
            }
        }
        if (sheetObj.worksheet.mergeCells) {
            sheetObj.worksheet.mergeCells = {
                "$t": "<% for(var m_cl=0; m_cl<_mergeCellArr_.length; m_cl++) { %><%-'<mergeCell ref=\"'+_mergeCellArr_[m_cl]+'\"/>'%><% } %>"
            };
        }
        if (sheetObj.worksheet.hyperlinks) {
            str = "<%for(var m_cl=0; m_cl<_hyperlinkArr_.length; m_cl++) { %><%-'<hyperlink ref=\"'+_hyperlinkArr_[m_cl].ref+'\"'%>";
            str += '<%var eny=_hyperlinkArr_[m_cl];var keyArr=Object.keys(eny);for(var tmp=0;tmp<keyArr.length;tmp++){var key=keyArr[tmp];if(key==="ref")continue;%><%-" "+key+"=\\""+eny[key]+"\\""%><%}%>';
            str += " /><%}%>";
            sheetObj.worksheet.hyperlinks = {
                "$t": str
            };
        }
        sheetBuf2 = Buffer.from(sheetSufStr.toString() + xml2json.toXml(sheetObj, "", {
            reSanitize: false
        }));
        reXmlEq = {
            reXmlEq: function (pixEq, jsStr, str) {
                // string
                if (pixEq === "=") {
                    jsStr =stripWhitespace(jsStr);
                    jsStr = `_ps_(${jsStr} ,buf)`;
                }
                // integer
                else if (pixEq === "~") {
                    jsStr =stripWhitespace(jsStr);
                    jsStr = `_pi_(${jsStr} ,buf)`;
                }
                // formula
                else if (pixEq === "#") {
                    jsStr =stripWhitespace(jsStr);
                    jsStr = `_pf_(${jsStr} ,buf)`;
                }
                return {
                    jsStr: jsStr,
                    str: str
                };
            }
        };
        reXmlEq.fileName = entry.fileName;
        str2 = ejs4xlx.parse(sheetBuf2, reXmlEq);
        str2 = "(async function anonymous(_args) {" + str2 + "})";
        anonymous = eval(str2);
        buffer2 = (await anonymous.call(this, data));
        if (entry.__remove_sheet) {
            await updateEntryAsync(entry.fileName);
        } else {
            buffer2 = Buffer.from(buffer2.toString().replace("<hyperlinks></hyperlinks>", "").replace("<mergeCells></mergeCells>", ""));
            await updateEntryAsync(entry.fileName, buffer2);
        }
    }
    sharedStrings2.push(Buffer.from("</sst>"));
    buffer2 = Buffer.concat(sharedStrings2);
    sharedStrings2 = void 0;
    await updateEntryAsync.apply(hzip, ["xl/sharedStrings.xml", buffer2]);
    await updateEntryAsync.apply(hzip, ["xl/workbook.xml", workbookBuf]);
    await updateEntryAsync.apply(hzip, ["xl/_rels/workbook.xml.rels", workbookRelsBuf]);
    return hzip.buffer;
};

module.exports={DataFactory,renderExcel};
