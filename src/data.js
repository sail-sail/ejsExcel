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
const _SHEET_STRING=[
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`, 
    `<%`,
        `var _data_ = _args._data_;`,
        `var _charPlus_ = _args._charPlus_.bind(_args);`,
        `var _charToNum_ = _args._charToNum_.bind(_args);`,
        `var _str2Xml_ = _args._str2Xml_.bind(_args);`,
        `var _hideSheet_ = _args._hideSheet_.bind(_args);`,
        `var _showSheet_ = _args._showSheet_.bind(_args);`,
        `var _deleteSheet_ = _args._deleteSheet_.bind(_args);`,
        `var _ps_ = _args._ps_.bind(_args);`,
        `var _pi_ = _args._pi_.bind(_args);`,
        `var _pf_ = _args._pf_.bind(_args);`,
        `var _acVar_ = _args._acVar_;`,
        `var _r = 0;`,
        `var _c = 0;`,
        `var _row = 0;`,
        `var _col = "";`,
        `var _rc = "";`,
        `var _img_ = _args._img_.bind(_args);`,
        `var _qrcode_ = _args._qrcode_.bind(_args);`,
        `var _mergeCellArr_ = [];`,
        `var _mergeCellFn_ = function(mclStr) {`,
        `	_mergeCellArr_.push(mclStr);`,
        `};`,
        `var _hyperlinkArr_ = [];`,
        `var _outlineLevel_ = _args._outlineLevel_.bind(_args);`,
    `%>`,
].join('\n');
const sheetSufStr = Buffer.from(_SHEET_STRING);

const _SHARED_STRING=[
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">`,
].join('\n');
const sharedStrings2Prx = Buffer.from(_SHARED_STRING);

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

/**
 * Data to render OpenXML 
 */
class Data{
    constructor(exlBuf,_data_){
        this.xjOp=xjOp;
        this._data_=_data_,
        this._acVar_ = {
            sharedStrings: [],
            _ss_len: 0,
        };
        this.sharedStrings2 = [sharedStrings2Prx];
        this.hzip=new Hzip(exlBuf);

        this.workbookEntry = this.hzip.getEntry("xl/workbook.xml");
        this.workbookRelsEntry =this.hzip.getEntry("xl/_rels/workbook.xml.rels");
        let entries = this.hzip.entries;
        this.sheetEntries = entries.filter(e=> e.fileName.indexOf("xl/worksheets/sheet") === 0 )
            .sort((arg0, arg1)=> arg0.fileName > arg1.fileName);
        this.sheetEntrieRels = entries.filter(e=> e.fileName.indexOf("xl/worksheets/_rels/") === 0)
            .sort((arg0, arg1) =>arg0.fileName > arg1.fileName);
        
        // a helper method
        this.updateEntryAsync= Promise_fromStandard(this.hzip.updateEntry,this.hzip);
    }

    _charPlus_(){
        return charPlus.apply(this,arguments);
    }
    _charToNum_(){
        return charToNum.apply(this,arguments);
    }
    _str2Xml_(){
        return str2Xml.apply(this,arguments);
    }

    /**
     * process string
     * @param {String} str 
     * @param {Buffer} buf 
     */
    _ps_(str, buf) {
        if (str === void 0) { str = ""; }     // undefined
        else if (str === null) { str = "NULL"; } // null
        str = str.toString();
        if (str === "") {
            var i,  l, ref2, tmpStr;
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
        let val = str2Xml(str);
        this.sharedStrings2.push(Buffer.from(`<si><t xml:space="preserve"> ${val} </t></si>`));
        let index = this._acVar_._ss_len;
        this._acVar_._ss_len++;
        return String(index);
    }
    
    /**
     * process formula
     * @param {String} str 
     * @param {Buffer} buf 
     */
    _pf_(str, buf) {
        if (str === void 0 || str === null) { str = ""; }
        str = str.toString();
        str = str2Xml(str);
        var i, l, m, ref2, ref3, tmpStr;
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
            puhStr = puhStr.toString();
            var index;
            if (puhStr.indexOf("</v>") !== -1) {
                index = Array.prototype.push.apply(buf, [Buffer.from(puhStr.replace(/<\/v>/m, "</f>"))]);
                buf.push = Array.prototype.push;
            } else {
                index = Array.prototype.push.apply(buf, [Buffer.from(puhStr)]);
            }
            return index;
        };
        return String(str);
    }


    /**
     * process integer
     * @param {*} str 
     * @param {*} buf 
     */
    _pi_(str, buf) {
        var i, l, ref2, tmpStr;
        if (str === true || str === false || isNaN(Number(str))) {
            return this._ps_(str, buf);
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
    }

    _outlineLevel_(str, buf) {
        let strNum = Number(str);
        if (isNaN(strNum) || strNum < 1) {
            return;
        }
        var i, l, ref2, tmpStr;
        for (i = l = ref2 = buf.length - 1; ref2 <= -1 ? l < -1 : l > -1; i = ref2 <= -1 ? ++l : --l) {
            tmpStr = buf[i].toString();
            if (/<row\s/gm.test(tmpStr)) {
                buf[i] = Buffer.from(replaceLast(tmpStr, /<row\s/gm, "<row outlineLevel=\"" + str + "\" "));
                break;
            }
        }
    }



    async _img_ (imgOpt, fileName, rowNum, cellNum) {

        if (isString(imgOpt) || Buffer.isBuffer(imgOpt)) {
            imgOpt = {
                imgPh: imgOpt
            };
        }
        imgOpt = imgOpt || {};
        if (!imgOpt.cellNumAdd) { imgOpt.cellNumAdd = 1; }
        if (!imgOpt.rowNumAdd) { imgOpt.rowNumAdd = 1; }
        imgOpt.cellNumAdd = Number(imgOpt.cellNumAdd) - 1;
        imgOpt.rowNumAdd = Number(imgOpt.rowNumAdd) - 1;

        const hzip=this.hzip;
        let imgBaseName = void 0;
        let imgBuf = void 0;
        let xdr_frt = imgOpt.xdr_frt;
        let imgPh = imgOpt.imgPh;
        if (!imgPh) { return ""; }
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
                let imgBufArr = [];
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
        let hashMd5 = crypto.createHash("md5");
        let md5Str = hashMd5.update(imgBuf).digest("hex");
        md5Str = "a" + md5Str;
        if (!imgBaseName) {
            imgBaseName = md5Str + ".png";
        }
        let cNvPrName = imgOpt.cNvPrName || imgBaseName;
        let cNvPrDescr = imgOpt.cNvPrDescr || imgBaseName;
        let cfileName = "xl/media/" + md5Str + ".png";
        let ref3 = this.hzip.entries;
        let itHs = false;
        let entryTmp=null;
        for (let m = 0; m < this.hzip.entries.length ; m++) {
            entryTmp =this.hzip.entries[m];
            if (entryTmp.fileName === cfileName) {
                itHs = true;
                break;
            }
        }
        if (!itHs) {
            await this.updateEntryAsync.apply(hzip, [cfileName, imgBuf, false]);
        }
        entryTmp = void 0;
        for (let n = 0; n <this.sheetEntries.length ; n++) {
            let sheetEntry = this.sheetEntries[n];
            if (fileName !== sheetEntry.fileName) {
                continue;
            }
            entryTmp = sheetEntry;
            break;
        }
        if (!entryTmp) {
            throw new Error(fileName + " dose not found!");
        }


        let doc= new DOMParser().parseFromString(
            (await inflateRawAsync(entryTmp.cfile)).toString(), 
            'text/xml'
        );
        let documentElement = doc.documentElement;
        let drawingEl = documentElement.getElementsByTagName("drawing")[0];
        if (!drawingEl) {
            throw "Excel模板显示动态图片之前,至少需要插入一张1像素的透明的图片,以初始化";
        }
        let drawingRId = drawingEl.getAttribute("r:id");
        let sheetEntrieRel = void 0;
        let sheetRelPth = path.dirname(fileName) + "/_rels/" + path.basename(fileName) + ".rels";
        for (let o = 0; o <this.sheetEntrieRels.length ; o++) {
            entryTmp = this.sheetEntrieRels[o];
            if (entryTmp.fileName === sheetRelPth) {
                sheetEntrieRel = entryTmp;
                break;
            }
        }
        if (!sheetEntrieRel) {
            throw new Error(sheetRelPth + " dose not found!");
        }
        doc = new DOMParser().parseFromString(
            (await inflateRawAsync(sheetEntrieRel.cfile)).toString(),
            'text/xml'
        );
        documentElement = doc.documentElement;
        let relationshipElArr = documentElement.getElementsByTagName("Relationship");
        let relationshipEl = void 0;
        for (let p = 0; p <relationshipElArr.length ; p++) {
            let shipEl = relationshipElArr[p];
            if (drawingRId === shipEl.getAttribute("Id")) {
                relationshipEl = shipEl;
                break;
            }
        }
        if (!relationshipEl) {
            throw new Error(sheetRelPth + (" Relationship.Id " + drawingRId + " dose not found!"));
        }
        let sei = path.basename(relationshipEl.getAttribute("Target")).replace("drawing", "").replace(".xml", "");
        sei = Number(sei) - 1;
        let drawingRelBuf = void 0;
        for (let q = 0; q <this.hzip.entries.length ; q++) {
            let entryImgTmp = this.hzip.entries[q];
            if (entryImgTmp.fileName === "xl/drawings/_rels/drawing" + (sei + 1) + ".xml.rels") {
                drawingRelBuf = (await inflateRawAsync(entryImgTmp.cfile));
                break;
            }
        }
        if (drawingRelBuf === void 0) {
            console.error("Excel模板显示动态图片之前,至少需要插入一张1像素的透明的图片,以初始化");
            return "";
        }
        let drawingRelObj = xml2json.toJson(drawingRelBuf, this.xjOp);
        if (!drawingRelObj["Relationships"]["Relationship"]) {
            drawingRelObj["Relationships"]["Relationship"] = [];
        } else if (!isArray(drawingRelObj["Relationships"]["Relationship"])) {
            drawingRelObj["Relationships"]["Relationship"] = [drawingRelObj["Relationships"]["Relationship"]];
        }
        itHs = false;
        let ref5 = drawingRelObj["Relationships"]["Relationship"];
        for (let r = 0; r < ref5.length ; r++) {
            let eny = ref5[r];
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
        let drawingRelStr = xml2json.toXml(drawingRelObj, "", {
            reSanitize: false
        });
        await this.updateEntryAsync.apply(this.hzip, ["xl/drawings/_rels/drawing" + (sei + 1) + ".xml.rels", Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + drawingRelStr)]);
        let drawingBuf = void 0;
        let ref6 = hzip.entries;
        for (let u = 0 ; u <ref6.length ; u++) {
            let entryImgTmp = ref6[u];
            if (entryImgTmp.fileName === "xl/drawings/drawing" + (sei + 1) + ".xml") {
                drawingBuf = (await inflateRawAsync(entryImgTmp.cfile));
                break;
            }
        }
        if (drawingBuf === void 0) {
            return "";
        }
        let drawingObj = xml2json.toJson(drawingBuf, this.xjOp);
        let TWO_CELL_ANCHOR=drawingObj["xdr:wsDr"]["xdr:twoCellAnchor"];
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

        let drawingStr = xml2json.toXml(drawingObj, "", {
            reSanitize: false
        });
        await this.updateEntryAsync.apply(this.hzip, ["xl/drawings/drawing" + (sei + 1) + ".xml", Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + drawingStr)]);
        return "";
    }

    async _qrcode_ (imgOpt, fileName, rowNum, cellNum) {
        if (!imgOpt || !imgOpt.text) {
            return "";
        }
        if (!imgOpt.margin) {
            imgOpt.margin = 0;
        }
        let qrBufArr = [];
        let imgSt = qr.image(imgOpt.text, imgOpt);
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
        let rvObj = (await this._img_(imgOpt, fileName, rowNum, cellNum));
        return rvObj;
    }


}


/**
 * 由于必须要创建一个Data对象用于渲染Excel，这里就简单用一个DataFactory函数预留一个接口，方便以后扩展Data，重写此方法
 * @param {*} exlBuf 
 * @param {*} _data_ 
 * @return {Data}
 */
function DataFactory(exlBuf,_data_){
    const d=new Data(exlBuf,_data_);
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


/**
 * 检查元素是否是Array，如果不是，转换为Array
 * @param {Array} arrayLike 
 */
function normalizeArray(arrayLike){
    if (!arrayLike) {
        arrayLike = [];
    } 
    else if (!isArray(arrayLike)) {
        arrayLike = [arrayLike];
    }
    return arrayLike;
}

/**
 * 在WorkbookRelsBuf中，找到元素的Target为filename的id
 * @param {Document} workbookRelsDoc 
 * @param {String} fileName 
 */
function findRelationshipByTarget(workbookRelsDoc,fileName){
    if (!fileName) { return null; }
    let documentElement = workbookRelsDoc.documentElement;
    let relationshipElArray = documentElement.getElementsByTagName("Relationship");
    let relationshipElement =null;
    
    for (let m = 0 ; m < relationshipElArray.length; m++) {
        relationshipElement = relationshipElArray[m];
        if ("xl/" + relationshipElement.getAttribute("Target") === fileName) {
            break;
        }
    }
    return relationshipElement;
}


function findSheetElementByRId(doc,rId){
    if(!doc || !rId){ return null; }
    let documentElement = doc.documentElement;
    let sheetsElement = documentElement.getElementsByTagName("sheets")[0];
    let sheetElementArray = sheetsElement.getElementsByTagName("sheet");
    let sheet = null;
    for (let n = 0; n <sheetElementArray.length; n++) {
        sheet = sheetElementArray[n];
        if (sheet.getAttribute("r:id") === rId) {
            break;
        }
    }
    return sheet;
}

function updateSheetHiddenState(workbookBuf,rId,state="hidden"||"show"){
    if (rId) {
        let doc = new DOMParser().parseFromString(workbookBuf.toString(), 'text/xml');
        let sheetElement=findSheetElementByRId(doc,rId);
        if(sheetElement){
            if(state=="hidden"){
                sheetElement.setAttribute("state", "hidden");
            }
            else{
                sheetElement.removeAttribute("state");
            }
        }
        workbookBuf = Buffer.from(doc.toString());
    }
    return workbookBuf;
}

const renderExcel = async function (exlBuf, _data_) {
    var  attr, attr0, attr_r, begin, buffer2, cEl, cElArr, cItem,  end, entry, hyperlink, hyperlinksDomEl, i, i1, idx, j1, key, keyArr,  len10, len11, len13, len14, len15, len3, len4, len5, len6, len7, len8, len9, m, m_c_i, mciNum, mciNumArr, mergeCell, n, o, p, pageMarginsDomEl, phoneticPr, phoneticPrDomEl, q, r,  ref, ref0, ref1, ref2, ref3, ref4, ref5, ref6, ref7, ref8, refArr, row, rowEl, rowElArr, sheetBuf, sheetBuf2,  si, si2, sirTp,  str, str2, u,  v, w,  x, xjOpTmp, y, z;
    let data = DataFactory(exlBuf,_data_);
    let sharedStrings2 = data.sharedStrings2;
    let sheetEntries = data.sheetEntries;
    let sheetEntrieRels =data.sheetEntrieRels;
    let workbookEntry = data.workbookEntry;
    let workbookRelsEntry =data.workbookRelsEntry; 
    let hzip =data.hzip;
    ref2 = hzip.entries;
    let updateEntryAsync = data.updateEntryAsync.bind(data);
    await updateEntryAsync("xl/calcChain.xml",Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></calcChain>'));
    let workbookBuf = (await inflateRawAsync(workbookEntry.cfile));
    let workbookRelsBuf = (await inflateRawAsync(workbookRelsEntry.cfile));


    data._hideSheet_=function(fileName){
        if (!workbookEntry || !workbookRelsEntry || !fileName) { return; }
        let doc = new DOMParser().parseFromString(workbookRelsBuf.toString(), 'text/xml');
        const target=findRelationshipByTarget(doc,fileName);
        const rId= !!target ?target.getAttribute('Id'):null;
        workbookBuf=updateSheetHiddenState(workbookBuf,rId,"hidden");
    }
    data._showSheet_ = function (fileName) {
        if (!workbookEntry || !workbookRelsEntry || !fileName) { return; }
        let doc = new DOMParser().parseFromString(workbookRelsBuf.toString(), 'text/xml');
        const target = findRelationshipByTarget(doc,fileName);
        const rId= !!target? target.getAttribute('Id'):null;
        workbookBuf=updateSheetHiddenState(workbookBuf,rId,"show");
    };
    data._deleteSheet_ = function (fileName) {
        if (!workbookEntry || !workbookRelsEntry || !fileName) {
            return;
        }
        let doc = new DOMParser().parseFromString(workbookRelsBuf.toString(), 'text/xml');
        let delRelationshipEl = findRelationshipByTarget(doc,fileName);
        let rId = !!delRelationshipEl? delRelationshipEl.getAttribute('Id'):null;
        if (delRelationshipEl) {
            doc.documentElement.removeChild(delRelationshipEl);
        }
        // update workbookRelEntry
        workbookRelsEntry = Buffer.from(doc.toString());

        doc = new DOMParser().parseFromString(workbookBuf.toString(), 'text/xml');
        let delSheet = findSheetElementByRId(doc,rId);

        if (delSheet) {
            let documentElement = doc.documentElement;
            let sheetsEl = documentElement.getElementsByTagName("sheets")[0];
            sheetsEl.removeChild(delSheet);
            let sheetElArr = sheetsEl.getElementsByTagName("sheet");
            let bookViewsEl = documentElement.getElementsByTagName("bookViews")[0];
            if (bookViewsEl) {
                let workbookViewEl = bookViewsEl.getElementsByTagName("workbookView")[0];
                if (workbookViewEl) {
                    let activeTab = workbookViewEl.getAttribute("activeTab");
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
        let sheetEntry=null;
        for (let o = 0; o < sheetEntries.length; o++) {
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
    normalizeArray(shsObj.sst.si);
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
        const startElement = xml2json.toJson.startElement;
        const endElement = xml2json.toJson.endElement;
        let sheetDataElementState = "";
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
            if (sheetDataElementState === "start") {
                return value;
            }
            return str2Xml(value) ;
        };
        let sheetObj = xml2json.toJson(sheetBuf, xjOpTmp);
        let SHEETDATA_ROW=sheetObj.worksheet.sheetData.row;
        if (SHEETDATA_ROW === void 0) { continue; } 
        else if (!isArray(SHEETDATA_ROW)) {
            sheetObj.worksheet.sheetData.row = [SHEETDATA_ROW];
        }
        let MERGE_CELLS=sheetObj.worksheet.mergeCells;
        if (MERGE_CELLS !== void 0 && MERGE_CELLS.mergeCell !== void 0) {
            MERGE_CELLS.mergeCell= normalizeArray(MERGE_CELLS.mergeCell);
        }
        let HYPER_LINKS=sheetObj.worksheet.hyperlinks;
        if(HYPER_LINKS && HYPER_LINKS.hyperlink){
            HYPER_LINKS.hyperlink = normalizeArray(HYPER_LINKS.hyperlink);
        }
        for (r = 0; r < SHEETDATA_ROW.length; r++) {
            let row = SHEETDATA_ROW[r];
            if (row.c !== void 0) {
                row.c=normalizeArray(row.c);
                for (u = 0; u < row.c.length ; u++) {
                    cItem = row.c[u];
                    // 共享字符串
                    if (cItem.t === "s" && cItem.v && !isNaN(Number(cItem.v["$t"])) && !cItem.f) {
                        si = shsObj.sst.si[cItem.v["$t"]];
                        phoneticPr = si.phoneticPr;
                        si2 = {
                            t: {
                                "$t": ""
                            }
                        };
                        if (si.r !== void 0) {
                            si.r=normalizeArray(si.r);
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
                    } 
                    else {
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
                    if (MERGE_CELLS !== void 0 && MERGE_CELLS.mergeCell !== void 0) {
                        mciNumArr = [];
                        ref8 = MERGE_CELLS.mergeCell;
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

                    if (HYPER_LINKS && HYPER_LINKS.hyperlink) {
                        mciNumArr = [];
                        for (m_c_i = y = 0; y < HYPER_LINKS.hyperlink.length; m_c_i = ++y) {
                            hyperlink = HYPER_LINKS.hyperlink[m_c_i];
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
                            HYPER_LINKS.hyperlink.splice(mciNum, 1);
                        }
                    }
                }
            }
        }
        if (MERGE_CELLS) {
            sheetObj.worksheet.mergeCells = {
                "$t": "<% for(var m_cl=0; m_cl<_mergeCellArr_.length; m_cl++) { %><%-'<mergeCell ref=\"'+_mergeCellArr_[m_cl]+'\"/>'%><% } %>"
            };
        }
        if (HYPER_LINKS) {
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

        let reXmlEq = {
            reXmlEq: function (pixEq, jsStr, str) {
                // echo string
                if (pixEq === "=") {
                    jsStr = stripWhitespace(jsStr);
                    jsStr = `_ps_(${jsStr} ,buf)`;
                } 
                // 数字类型
                else if (pixEq === "~") {
                    jsStr = stripWhitespace(jsStr);
                    jsStr = `_pi_(${jsStr} ,buf)`;
                } 
                // 动态公式
                else if (pixEq === "#") {
                    jsStr = stripWhitespace(jsStr);
                    jsStr = `_pf_(${jsStr} ,buf)`;
                }
                return {
                    jsStr: jsStr,
                    str: str
                };
            }
        };
        reXmlEq.fileName = entry.fileName;
        str2 = ejs4xlx.parse(sheetBuf2, reXmlEq);   // 解析出代码片段
        str2 = `(async function anonymous(_args) { ${str2} })`;  // 构造函数字符串
        // dangerous  
        let anonymous = eval(str2);
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
