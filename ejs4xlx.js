
/*!
 * EJS
 * Copyright(c) 2012 TJ Holowaychuk <tj@vision-media.ca>
 * MIT Licensed
 */

/**
 * Module dependencies.
 */

var fs = require('fs');
var path = require('path');
var basename = path.basename;
var dirname = path.dirname;
var extname = path.extname;
var join = path.join;

var isType = function(type) {
	return function(obj) {
		return Object.prototype.toString.call(obj) === "[object " + type + "]";
	};
};
var isObject = isType("Object");
var isString = isType("String");
var isArray = Array.isArray || isType("Array");
var isFunction = isType("Function");
var UID = Date.now()
var uniqueID = function() {
	return (UID++).toString(36);
}
var charToNum = function(str) {
  var i, j, len, m, ref, temp, val;
  str = new String(str);
  val = 0;
  len = str.length;
  for (j = m = 0, ref = len; 0 <= ref ? m < ref : m > ref; j = 0 <= ref ? ++m : --m) {
    i = len - 1 - j;
    temp = str.charCodeAt(i) - 65 + 1;
    val += temp * Math.pow(26, j);
  }
  return val;
};

/**
 * Filters.
 * 
 * @type Object
 */

var filters = exports.filters = require('./filters');

/**
 * Translate filtered code into function calls.
 *
 * @param {String} js
 * @return {String}
 * @api private
 */

function filtered(js) {
  return js.substr(1).split('|').reduce(function(js, filter){
    var parts = filter.split(':')
      , name = parts.shift()
      , args = parts.shift() || '';
    if (args) args = ', ' + args;
    return 'filters.' + name + '(' + js + args + ')';
  });
};

/**
 * Re-throw the given `err` in context to the
 * `str` of ejs, `filename`, and `lineno`.
 *
 * @param {Error} err
 * @param {String} str
 * @param {String} filename
 * @param {String} lineno
 * @api private
 */

function rethrow(err, str, filename, lineno){
  var lines = str.split('\n')
    , start = Math.max(lineno - 3, 0)
    , end = Math.min(lines.length, lineno + 3);

  // Error context
  var context = lines.slice(start, end).map(function(line, i){
    var curr = i + start + 1;
    return (curr == lineno ? ' >> ' : '    ')
      + curr
      + '| '
      + line;
  }).join('\n');

  // Alter exception message
  err.path = filename;
  err.message = (filename || 'ejs') + ':' 
    + lineno + '\n' 
    + context + '\n\n' 
    + err.message;
  
  throw err;
}

/**
 * Parse the given `str` of ejs, returning the function body.
 *
 * @param {String} str
 * @return {String}
 * @api public
 */

var parse = exports.parse = function(str, options){
	//sail 2013-04-21
	if(Buffer.isBuffer(str)) str = str.toString();
	options = options || {};
  var open = options.open || exports.open || '<%'
    , close = options.close || exports.close || '%>'
    , beforeBuf = options.beforeBuf || exports.beforeBuf || ''
    , filename = options.filename
    //=,- 符号时生成的字符串针对xml转义符进行转义,在ejsExcel中会使用到
    , reXmlEq = options.reXmlEq
    , buf = [];
  buf.push('var buf = [];');
  buf.push(beforeBuf);
  buf.push(';buf.push(new Buffer(\'');

  var lineno = 1;

  //sail
  var isForRowBegin = false;
  var isForRowEnd = false;
  var isForCellBegin = false;
  var isForCellEnd = false;
  //行号
  var rowRn = undefined;
  //列号
  var cellRn = undefined;
  var isCEnd = false;
  var isCbegin = false;
  var isRowBegin = false;
  var isRowEnd = false;
  var consumeEOL = false;
  
  var forRBegin = false;
  var forREnd = false;
  var forRNumArr = [];
  var forCBegin = false;
  var forCEnd = false;
  var forCNumArr = [];
  var ifRBegin = false;
  var ifREnd = false;
  var ifCBegin = false;
  var ifCEnd = false;
  var pixEq = "";
  for (var i = 0, len = str.length; i < len; ++i) {
    if (str.slice(i, open.length + i) == open) {
      i += open.length;
      
      //sail
      pixEq = str.substr(i, 1);
      var prefix, postfix, line = lineno;
      if(pixEq === "=" || pixEq === "-" || pixEq === "~" || pixEq === "#") {
    	  prefix = "'));buf.push(new Buffer(";
          postfix = "));buf.push(new Buffer('";
          ++i;
      } else {
    	  prefix = "'));";
          postfix = ";buf.push(new Buffer('";
      }

      var end = str.indexOf(close, i)
        , js = str.substring(i, end)
        , start = i
        , include = null
        , n = 0;

      if ('-' == js[js.length-1]){
        js = js.substring(0, js.length - 2);
        consumeEOL = true;
      }
      
      js = js.trim();
      if(0 == js.indexOf('forRow')) {
    	  var uid = uniqueID();
    	  var name = js.slice(6);
    	  isForRowBegin = true;
    	  isForRowEnd = false;
    	  var nameArr = [];
    	  nameArr[0] = name.substring(0,name.indexOf(" in "));
    	  nameArr[1] = name.substring(name.indexOf(" in ")+4);
    	  var pixJs = "";
    	  if(nameArr[1].indexOf("|||") !== -1) {
    		  pixJs = nameArr[1].substring(nameArr[1].indexOf("|||")+3);
    		  nameArr[1] = nameArr[1].substring(0,nameArr[1].indexOf("|||"));
    	  }
    	  var itemName = nameArr[0].trim();
    	  var iName = "";
    	  if(itemName.indexOf(",") !== -1) {
    		  var tmpArr = itemName.split(",");
    		  itemName = tmpArr[0].trim();
    		  if(tmpArr[1].trim() !== "") {
    			  iName = "var "+tmpArr[1].trim()+"=I_m"+uid+";";
    		  }
    	  }
    	  var arrName = nameArr[1].trim();
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<row r="/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<row r="/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "'));var I_rLen"+uid+"="+arrName+";if(Array.isArray(I_rLen"+uid+")){I_rLen"+uid+"=I_rLen"+uid+".length;} for(var I_m"+uid+"=0;I_m"+uid+"<I_rLen"+uid+";I_m"+uid+"++){yield Promise_sleep(0);"+iName+"var "+itemName+"="+arrName+"[I_m"+uid+"];if(typeof("+arrName+")===\"number\"){"+itemName+"=I_m"+uid+";}"+pixJs+";buf.push(new Buffer('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      //sail 2015-06-24 --begin
      else if(0 == js.indexOf('ifRBegin')) {
    	  ifRBegin = true;
    	  ifREnd = false;
    	  forRNumArr.push(rowRn);
    	  var jsStr = js.slice(8);
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<row r="/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<row r="/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "'));if("+jsStr+"){buf.push(new Buffer('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      } else if(0 == js.indexOf('ifREnd')) {
    	  ifRBegin = false;
    	  ifREnd = true;
    	  if(forRNumArr.length === 0) throw new Error("ifREnd must begin with ifRBegin");
    	  var rjsNum = rowRn-forRNumArr.pop();
    	  var rjsStr = js.slice(6);
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<\/row>/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<\/row>/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "</row>'));_r+="+rjsNum+";}_r-="+rjsNum+";buf.push(new Buffer('";
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      else if(0 == js.indexOf('ifCBegin')) {
    	  ifCBegin = true;
    	  ifCEnd = false;
    	  forCNumArr.push(cellRn);
    	  var jsStr = js.slice(8);
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<c r="/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<c r="/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "'));if("+jsStr+"){buf.push(new Buffer('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      } else if(0 == js.indexOf('ifCEnd')) {
    	  ifCBegin = false;
    	  ifCEnd = true;
    	  if(forCNumArr.length === 0) throw new Error("ifCEnd must be begin with ifCBegin");
    	  var rjsNum = charToNum(cellRn)-charToNum(forCNumArr.pop());
    	  var rjsStr = js.slice(6);
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<\/c>/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<\/c>/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "</c>'));_c+="+rjsNum+";}_c-="+rjsNum+";buf.push(new Buffer('";
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      else if(0 == js.indexOf('forCBegin')) {
    	  forCBegin = true;
    	  forCEnd = false;
    	  forCNumArr.push(cellRn);
    	  var uid = uniqueID();
    	  var name = js.slice(9);
    	  var nameArr = [];
    	  nameArr[0] = name.substring(0,name.indexOf(" in "));
    	  nameArr[1] = name.substring(name.indexOf(" in ")+4);
    	  var pixJs = "";
    	  if(nameArr[1].indexOf("|||") !== -1) {
    		  pixJs = nameArr[1].substring(nameArr[1].indexOf("|||")+3);
    		  nameArr[1] = nameArr[1].substring(0,nameArr[1].indexOf("|||"));
    	  }
    	  var itemName = nameArr[0].trim();
    	  var iName = "";
    	  if(itemName.indexOf(",") !== -1) {
    		  var tmpArr = itemName.split(",");
    		  itemName = tmpArr[0].trim();
    		  if(tmpArr[1].trim() !== "") {
    			  iName = "var "+tmpArr[1].trim()+"=I_c"+uid+";";
    		  }
    	  }
    	  var arrName = nameArr[1].trim();
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<c r="/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<c r="/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "'));var I_cLen"+uid+"="+arrName+";if(Array.isArray(I_cLen"+uid+")){I_cLen"+uid+"=I_cLen"+uid+".length;} for(var I_c"+uid+"=0;I_c"+uid+"<I_cLen"+uid+";I_c"+uid+"++){yield Promise_sleep(0);"+iName+"var "+itemName+"="+arrName+"[I_c"+uid+"];if(typeof("+arrName+")===\"number\"){"+itemName+"=I_c"+uid+";}"+pixJs+";buf.push(new Buffer('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
    	  
      } else if(0 == js.indexOf('forCEnd')) {
    	  forCBegin = false;
    	  forCEnd = true;
    	  if(forCNumArr.length === 0) throw new Error("forCEnd must be begin with forCBegin");
    	  var rjsNum = charToNum(cellRn)-charToNum(forCNumArr.pop());
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<\/c>/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<\/c>/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "</c>'));_c+="+rjsNum+";}_c-="+rjsNum+";buf.push(new Buffer('";
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      //sail 2015-06-24 --end
      //sail 2014-04-09 --begin
      else if(0 == js.indexOf('forRBegin')) {
    	  forRBegin = true;
    	  forREnd = false;
    	  forRNumArr.push(rowRn);
    	  var uid = uniqueID();
    	  var name = js.slice(9);
    	  var nameArr = [];
    	  nameArr[0] = name.substring(0,name.indexOf(" in "));
    	  nameArr[1] = name.substring(name.indexOf(" in ")+4);
    	  var pixJs = "";
    	  if(nameArr[1].indexOf("|||") !== -1) {
    		  pixJs = nameArr[1].substring(nameArr[1].indexOf("|||")+3);
    		  nameArr[1] = nameArr[1].substring(0,nameArr[1].indexOf("|||"));
    	  }
    	  var itemName = nameArr[0].trim();
    	  var iName = "";
    	  if(itemName.indexOf(",") !== -1) {
    		  var tmpArr = itemName.split(",");
    		  itemName = tmpArr[0].trim();
    		  if(tmpArr[1].trim() !== "") {
    			  iName = "var "+tmpArr[1].trim()+"=I_m"+uid+";";
    		  }
    	  }
    	  var arrName = nameArr[1].trim();
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<row r="/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<row r="/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "'));var I_rLen"+uid+"="+arrName+";if(Array.isArray(I_rLen"+uid+")){I_rLen"+uid+"=I_rLen"+uid+".length;} for(var I_m"+uid+"=0;I_m"+uid+"<I_rLen"+uid+";I_m"+uid+"++){yield Promise_sleep(0);"+iName+"var "+itemName+"="+arrName+"[I_m"+uid+"];if(typeof("+arrName+")===\"number\"){"+itemName+"=I_m"+uid+";}"+pixJs+";buf.push(new Buffer('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      else if(0 == js.indexOf('forREnd')) {
    	  forRBegin = false;
    	  forREnd = true;
    	  if(forRNumArr.length === 0) throw new Error("forREnd must begin with forRBegin");
    	  var rjsNum = rowRn-forRNumArr.pop();
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<\/row>/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<\/row>/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "</row>'));_r+="+rjsNum+";}_r-="+rjsNum+";buf.push(new Buffer('";
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      //sail 2014-04-09 --end
      else if(0 === js.indexOf('forCell')) {
    	  var uid = uniqueID();
    	  var name = js.slice(7);
    	  isForCellBegin = true;
    	  isForCellEnd = false;
    	  var nameArr = [];
    	  nameArr[0] = name.substring(0,name.indexOf(" in "));
    	  nameArr[1] = name.substring(name.indexOf(" in ")+4);
    	  var pixJs = "";
    	  if(nameArr[1].indexOf("|||") !== -1) {
    		  pixJs = nameArr[1].substring(nameArr[1].indexOf("|||")+3);
    		  nameArr[1] = nameArr[1].substring(0,nameArr[1].indexOf("|||"));
    	  }
    	  var itemName = nameArr[0].trim();
    	  var iName = "";
    	  if(itemName.indexOf(",") !== -1) {
    		  var tmpArr = itemName.split(",");
    		  itemName = tmpArr[0].trim();
    		  if(tmpArr[1].trim() !== "") {
    			  iName = "var "+tmpArr[1].trim()+"=I_c"+uid+";";
    		  }
    	  }
    	  var arrName = nameArr[1].trim();
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<c r="/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<c r="/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "'));var I_cLen"+uid+" = "+arrName+";if(Array.isArray(I_cLen"+uid+")){I_cLen"+uid+"=I_cLen"+uid+".length;} for(var I_c"+uid+"=0;I_c"+uid+"<I_cLen"+uid+";I_c"+uid+"++){yield Promise_sleep(0);"+iName+"var "+itemName+"="+arrName+"[I_c"+uid+"];if(typeof("+arrName+")===\"number\"){"+itemName+"=I_c"+uid+";}"+pixJs+";buf.push(new Buffer('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      //图片
      else if(0 == js.indexOf('_img_(')) {
    	  if(options !== undefined && options.fileName !== undefined) {
    		  js = js.substring(0,js.length-1);
    		  var cellNum = 1;
    		  for(var sei=0; sei<cellRn.length; sei++) {
    			  cellNum += cellRn.charCodeAt(sei)-65+(cellRn.length-1-sei)*26;
    		  }
    		  js = "yield "+js+",\""+options.fileName.replace(/\"/gm,"\\\"")+"\",("+rowRn+"+_r),("+cellNum+"+_c))";
    	  }
      }
      //二维码图片
      else if(0 == js.indexOf('_qrcode_(')) {
    	  if(options !== undefined && options.fileName !== undefined) {
    		  js = js.substring(0,js.length-1);
    		  var cellNum = 1;
    		  for(var sei=0; sei<cellRn.length; sei++) {
    			  cellNum += cellRn.charCodeAt(sei)-65+(cellRn.length-1-sei)*26;
    		  }
    		  js = "yield "+js+",\""+options.fileName.replace(/\"/gm,"\\\"")+"\",("+rowRn+"+_r),("+cellNum+"+_c))";
    	  }
      }
      //分组
      else if(0 === js.indexOf('_outlineLevel_(')) {
    	  js = js.substring(0,js.length-1);
    	  js += ",buf)";
      }

      while (~(n = js.indexOf("\n", n))) n++, lineno++;
      if (js.substr(0, 1) == ':') js = filtered(js);
      if (js) {
        if (js.lastIndexOf('//') > js.lastIndexOf('\n')) js += '\n';
        //sail 2012-04-21
        //=,- 符号时生成的字符串针对xml转义符进行转义,在ejsExcel中会使用到
        if(reXmlEq !== undefined && (pixEq === "=" || pixEq === "~" || pixEq === "#")) {
        	var rxeObj = reXmlEq(pixEq,js,buf.join(''));
        	js = rxeObj.jsStr;
        	buf = [rxeObj.str];
        }
        //sail 2012-04-21 end
        buf.push(prefix, js, postfix);
      }
      i += end - start + close.length - 1;

    }
    else if(str.substr(i,8) === "<row r=\"") {
    	isRowBegin = true;
  		isRowEnd = false;
  		forCNumArr = []
  		i += 7;
    	buf.push("<row r=\"");
    }
    else if(isRowBegin === true && isRowEnd === false && str.substr(i,1) === "\"") {
    	isRowEnd = true;
    	isRowBegin = false;
    	var strTmp = buf.join('')+"\"";
  	    var mthArr = strTmp.match(/<row r="\d+"/gm);
  	    var mthLt = mthArr[mthArr.length-1];
  	    //行号
  	    rowRn = mthLt.replace(/<row r="/gm,"").replace(/"/gm,"");
  	    var repNum = 0;
  	    strTmp = strTmp.replace(/<row r="\d+"/gm,function(s){
  	  		repNum++;
  	  		if(mthArr.length === repNum) {
  	  			return "<row r=\"'));_c=0;_row="+rowRn+"+_r;buf.push(new Buffer(String(_row)));buf.push(new Buffer('\"";
  	  		}
  	  		return s;
  	    });
  	    buf = [strTmp];
    }
    else if(str.substr(i,6) === "<c r=\"") {
    	isCbegin = true;
    	isCEnd = false;
    	i += 5;
    	buf.push("<c r=\"");
    }
    else if(isCbegin === true && isCEnd === false && str.substr(i,1) === "\"") {
    	isCEnd = true;
    	isCbegin = false;
    	var strTmp = buf.join('')+"\"";
    	var mthArr = strTmp.match(/<c r="[A-Z]+[0-9]+"/gm);
    	var mthLt = mthArr[mthArr.length-1];
  		cellRn = mthLt.replace(/<c r="/gm,"").replace(/\d+"/gm,"");
  		var repNum = 0;
  		strTmp = strTmp.replace(/<c r="[A-Z]+[0-9]+"/gm,function(s){
  	  	  repNum++;
  	  	  if(mthArr.length === repNum) {
  	  		  return "<c r=\"'));_col=_charPlus_('"+cellRn+"',_c);_rc=_col+_row;buf.push(new Buffer(_rc));buf.push(new Buffer('\"";
  	  	  }
  	  	  return s;
  	    });
  	    buf = [strTmp];
    }
    //sail
    else if(isForRowBegin === true && isForRowEnd === false && str.substr(i,6) === "</row>") {
    	isForRowEnd = true;
    	isForRowBegin = false;
    	i += 5;
    	buf.push("</row>'));_r++;}_r--;buf.push(new Buffer('");
    }
    else if(isForCellBegin === true && isForCellEnd === false && str.substr(i,4) === "</c>") {
    	isForCellEnd = true;
    	isForCellBegin = false;
    	i += 3;
    	buf.push("</c>'));_c++;}_c--;buf.push(new Buffer('");
    }
    else if (str.substr(i, 1) == "\\") {
      buf.push("\\\\");
    } else if (str.substr(i, 1) == "'") {
      buf.push("\\'");
    } else if (str.substr(i, 1) == "\r") {
      buf.push(" ");
    } else if (str.substr(i, 1) == "\n") {
      if (consumeEOL) {
        consumeEOL = false;
      } else {
        buf.push("\\n");
        lineno++;
      }
    } else {
      pixEq = "";
      buf.push(str.substr(i, 1));
    }
  }

  buf.push("'));");
  buf.push("buf = Buffer.concat(buf);return buf;");
  //fs.writeFileSync("C:/abc.js",buf.join(''));
  //process.exit();
  return buf.join('');
};

/**
 * Resolve include `name` relative to `filename`.
 *
 * @param {String} name
 * @param {String} filename
 * @return {String}
 * @api private
 */

function resolveInclude(name, filename) {
  var path = join(dirname(filename), name);
  var ext = extname(name);
  if (!ext) path += '.ejs';
  return path;
}
