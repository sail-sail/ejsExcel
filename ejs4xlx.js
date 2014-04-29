
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
  buf.push('\n buf.push(\'');

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
  var pixEq = "";
  for (var i = 0, len = str.length; i < len; ++i) {
    if (str.slice(i, open.length + i) == open) {
      i += open.length;
      
      //sail
      pixEq = str.substr(i, 1);
      var prefix, postfix, line = lineno;
      if(pixEq === "=" || pixEq === "-" || pixEq === "~" || pixEq === "#") {
    	  prefix = "');buf.push(";
          postfix = ");buf.push('";
          ++i;
      } else {
    	  prefix = "');";
          postfix = ";buf.push('";
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
      if (0 == js.indexOf('include')) {
        var name = js.slice(7);
        if (!filename) throw new Error('filename option is required for includes');
        var path = resolveInclude(name, filename);
        include = fs.readSync(path, 'utf8');
        extname = path.extname(path);
        if(extname === "ejs") {
        	include = exports.parse(include, { filename: path, open: open, close: close });
        	buf.push("' + (function(){" + include + "})() + '");
        } else {
        	buf.push("'" + include + " '");
        }
        js = '';
      } else if(0 == js.indexOf('forRow')) {
    	  var name = js.slice(6);
    	  isForRowBegin = true;
    	  isForRowEnd = false;
    	  var nameArr = [];
    	  nameArr[0] = name.substring(0,name.indexOf(" in "));
    	  //var nameArr = name.split(" in ");
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
    			  iName = "var "+tmpArr[1].trim()+"=I_m;";
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
    			  return "');var I_rLen = "+arrName+";if(Array.isArray("+arrName+")){I_rLen=("+arrName+").length;};for(var I_m=0;I_m<I_rLen;I_m++){$await(Wind.Async.sleep(0));"+iName+"var "+itemName+"="+arrName+"[I_m];if(typeof("+arrName+")===\"number\"){"+itemName+"=I_m;}"+pixJs+";buf.push('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      //sail 2014-04-09 --begin
      else if(0 == js.indexOf('forRBegin')) {
    	  var name = js.slice(9);
    	  var nameArr = [];
    	  nameArr[0] = name.substring(0,name.indexOf(" in "));
    	  //var nameArr = name.split(" in ");
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
    			  iName = "var "+tmpArr[1].trim()+"=I_m;";
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
    			  return "');var I_rLen = "+arrName+";if(Array.isArray("+arrName+")){I_rLen=("+arrName+").length;};for(var I_m=0;I_m<I_rLen;I_m++){$await(Wind.Async.sleep(0));"+iName+"var "+itemName+"="+arrName+"[I_m];if(typeof("+arrName+")===\"number\"){"+itemName+"=I_m;}"+pixJs+";buf.push('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      else if(0 == js.indexOf('forREnd')) {
    	  var rjsStr = js.slice(7);
    	  var strTmp = buf.join('');
    	  var mthArr = strTmp.match(/<row r="/gm);
    	  var mthLt = mthArr[mthArr.length-1];
    	  var repNum = 0;
    	  strTmp = strTmp.replace(/<row r="/gm,function(s){
    		  repNum++;
    		  if(mthArr.length === repNum) {
    			  return "');_r+=Number(eval('"+rjsStr+"'));}_r-=Number(eval('"+rjsStr+"'));buf.push('"+mthLt;
    		  }
    		  return s;
    	  });
    	  buf = [strTmp];
    	  js = '';
      }
      //sail 2014-04-09 --end
      else if(0 === js.indexOf('forCell')) {
    	  var name = js.slice(7);
    	  isForCellBegin = true;
    	  isForCellEnd = false;
    	  var nameArr = [];
    	  nameArr[0] = name.substring(0,name.indexOf(" in "));
    	  //var nameArr = name.split(" in ");
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
    			  iName = "var "+tmpArr[1].trim()+"=I_c;";
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
    			  return "');var I_cLen = "+arrName+";if(Array.isArray("+arrName+")){I_cLen=("+arrName+").length;};for(var I_c=0;I_c<I_cLen;I_c++){$await(Wind.Async.sleep(0));"+iName+"var "+itemName+"="+arrName+"[I_c];if(typeof("+arrName+")===\"number\"){"+itemName+"=I_c;}"+pixJs+";buf.push('"+mthLt;
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
    		  js = "$await("+js+",\""+options.fileName.replace(/\"/gm,"\\\"")+"\","+rowRn+","+cellNum+"))";
    	  }
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
  		i += 7;
    	buf.push("<row r=\"");
    }
    else if(isRowBegin === true && isRowEnd === false && str.substr(i,1) === "\"") {
    	isRowEnd = true;
    	isRowBegin = false;
    	var strTmp = buf.join('')+"\"";
  	    var mthArr = strTmp.match(/<row\s+r="\d+"/gm);
  	    var mthLt = mthArr[mthArr.length-1];
  	    //行号
  	    rowRn = mthLt.replace(/<row\s+r="/gm,"").replace(/"/gm,"");
  	    var repNum = 0;
  	    strTmp = strTmp.replace(/<row\s+r="\d+"/gm,function(s){
  	  		repNum++;
  	  		if(mthArr.length === repNum) {
  	  			return "<row r=\"');_c=0;buf.push("+rowRn+"+_r);buf.push('\"";
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
    	var mthArr = strTmp.match(/<c\s+r="\D+\d+"/gm);
    	var mthLt = mthArr[mthArr.length-1];
  		cellRn = mthLt.replace(/<c\sr="/gm,"").replace(/\d+"/gm,"");
  		var repNum = 0;
  		strTmp = strTmp.replace(/<c\s+r="\D+\d+"/gm,function(s){
  	  	  repNum++;
  	  	  if(mthArr.length === repNum) {
  	  		  return "<c r=\"');buf.push(_charPlus_('"+cellRn+"',_c));buf.push("+rowRn+"+_r);buf.push('\"";
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
    	buf.push("</row>');_r++;}_r--;buf.push('");
    }
    else if(isForCellBegin === true && isForCellEnd === false && str.substr(i,4) === "</c>") {
    	isForCellEnd = true;
    	isForCellBegin = false;
    	i += 3;
    	buf.push("</c>');_c++;}_c--;buf.push('");
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

  buf.push("');\nreturn buf.join('');");
  //fs.writeFileSync("C:/abc.js",buf.join(''));
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
