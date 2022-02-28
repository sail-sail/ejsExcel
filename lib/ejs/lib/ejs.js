
/*!
 * EJS
 * Copyright(c) 2012 TJ Holowaychuk <tj@vision-media.ca>
 * MIT Licensed
 */

/**
 * Filters.
 * 
 * @type Object
 */

var filters = { };
exports.filters = filters;

filters.first = function(obj) {
  return obj[0];
};

filters.last = function(obj) {
  return obj[obj.length - 1];
};

filters.capitalize = function(str){
  str = String(str);
  return str[0].toUpperCase() + str.substr(1, str.length);
};

filters.downcase = function(str){
  return String(str).toLowerCase();
};

filters.upcase = function(str){
  return String(str).toUpperCase();
};

filters.sort = function(obj){
  return Object.create(obj).sort();
};

/**
 * Sort the target `obj` by the given `prop` ascending.
 */

filters.sort_by = function(obj, prop){
  return Object.create(obj).sort(function(a, b){
    a = a[prop], b = b[prop];
    if (a > b) return 1;
    if (a < b) return -1;
    return 0;
  });
};

/**
 * Size or length of the target `obj`.
 */

filters.size = function(obj) {
  return obj.length;
};

/**
 * Add `a` and `b`.
 */

filters.plus = function(a, b){
  return Number(a) + Number(b);
};

/**
 * Subtract `b` from `a`.
 */

filters.minus = function(a, b){
  return Number(a) - Number(b);
};

/**
 * Multiply `a` by `b`.
 */

filters.times = function(a, b){
  return Number(a) * Number(b);
};

/**
 * Divide `a` by `b`.
 */

filters.divided_by = function(a, b){
  return Number(a) / Number(b);
};

/**
 * Join `obj` with the given `str`.
 */

filters.join = function(obj, str){
  return obj.join(str || ', ');
};

/**
 * Truncate `str` to `len`.
 */

filters.truncate = function(str, len){
  str = String(str);
  return str.substr(0, len);
};

/**
 * Truncate `str` to `n` words.
 */

filters.truncate_words = function(str, n){
  var str = String(str)
    , words = str.split(/ +/);
  return words.slice(0, n).join(' ');
};

/**
 * Replace `pattern` with `substitution` in `str`.
 */

filters.replace = function(str, pattern, substitution){
  return String(str).replace(pattern, substitution || '');
};

/**
 * Prepend `val` to `obj`.
 */

filters.prepend = function(obj, val){
  return Array.isArray(obj)
    ? [val].concat(obj)
    : val + obj;
};

/**
 * Append `val` to `obj`.
 */

filters.append = function(obj, val){
  return Array.isArray(obj)
    ? obj.concat(val)
    : obj + val;
};

/**
 * Map the given `prop`.
 */

filters.map = function(arr, prop){
  return arr.map(function(obj){
    return obj[prop];
  });
};

/**
 * Reverse the given `obj`.
 */

filters.reverse = function(obj){
  return Array.isArray(obj)
    ? obj.reverse()
    : String(obj).split('').reverse().join('');
};

/**
 * Get `prop` of the given `obj`.
 */

filters.get = function(obj, prop){
  return obj[prop];
};

/**
 * Packs the given `obj` into json string
 */
filters.json = function(obj){
  return JSON.stringify(obj);
};

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

var parse = function(str, options) {
	options = options || {};
  var open = options.open || '<%'
    , close = options.close || '%>'
    , filename = options.filename
    //=,- 符号时生成的字符串针对xml转义符进行转义,在ejsExcel中会使用到
    , reXmlEq = options.reXmlEq
    , buf = [];
  buf.push('var buf = \'\';');
  buf.push(';buf += \'');

  var lineno = 1;

  var consumeEOL = false;
  for (var i = 0, len = str.length; i < len; ++i) {
    if (str.slice(i, open.length + i) == open) {
      i += open.length
      
      //sail
      var pixEq = str.substr(i, 1);
      var prefix, postfix, line = lineno;
      if(pixEq === "=" || pixEq === "-" || pixEq === "~" || pixEq === "#") {
    	  prefix = "';buf += String(";
          postfix = ");buf += '";
          ++i;
      } else {
    	  prefix = "';";
          postfix = ";buf += '";
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

  buf.push("';");
  buf.push("return buf;");
  return buf.join('');
};
exports.parse = parse;
