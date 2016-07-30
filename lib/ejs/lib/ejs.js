
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

  var consumeEOL = false;
  for (var i = 0, len = str.length; i < len; ++i) {
    if (str.slice(i, open.length + i) == open) {
      i += open.length
      
      //sail
      var pixEq = str.substr(i, 1);
      var prefix, postfix, line = lineno;
      if(pixEq === "=" || pixEq === "-" || pixEq === "~" || pixEq === "#") {
    	  prefix = "'));buf.push(new Buffer(String(";
          postfix = ")));buf.push(new Buffer('";
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

      while (~(n = js.indexOf("\n", n))) n++, lineno++;
      if (js.substr(0, 1) == ':') js = filtered(js);
      if (js) {
        if (js.lastIndexOf('//') > js.lastIndexOf('\n')) js += '\n';
        //sail 2012-04-21
        //=,- 符号时生成的字符串针对xml转义符进行转义,在ejsExcel中会使用到
        if(reXmlEq !== undefined && (pixEq === "=" || pixEq === "-" || pixEq === "~" || pixEq === "#")) js = reXmlEq(pixEq,js);
        //sail 2012-04-21 end
        buf.push(prefix, js, postfix);
      }
      i += end - start + close.length - 1;

    } else if (str.substr(i, 1) == "\\") {
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
      buf.push(str.substr(i, 1));
    }
  }

  buf.push("'));");
  buf.push("buf=Buffer.concat(buf);return buf;");
  //fs.writeFileSync("C:/abc2.js",buf.join(''));
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
