
/*!
 * EJS
 * Copyright(c) 2012 TJ Holowaychuk <tj@vision-media.ca>
 * MIT Licensed
 */

/**
 * Module dependencies.
 */

var Wind = require("../../Wind");
var Binding = Wind.Async.Binding;
var fs = require('fs');
var readAsync = Binding.fromStandard(fs.readFile);
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

var parse = exports.parse = (function (str, options) {
    var _builder_$0 = Wind.b["async"];
    var _arguments_$ = arguments;
    var _caller_$0 = this.$caller;
    return _builder_$0.m(this,
        _builder_$0.e(function() {
            this.$caller = _caller_$0;
            if (Buffer.isBuffer(str)) 
                str = str.toString();
            options = options || { };
            var open = options.open || exports.open || "<%", close = options.close || exports.close || "%>", beforeBuf = options.beforeBuf || exports.beforeBuf || "", filename = options.filename, reXmlEq = options.reXmlEq, buf = [];
            buf.push("var buf = [];");
            buf.push(beforeBuf);
            buf.push("\n buf.push('");
            var lineno = 1;
            var consumeEOL = false;
            return _builder_$0.f(
                _builder_$0.e(function() {
                    var i = 0, len = str.length;
                    return _builder_$0.a(function() {
                        return i < len;
                    }, function() {
                        ++ i;
                    },
                        _builder_$0.e(function() {
                            return _builder_$0.n(Wind.Async.sleep(0), function () {
                                if (str.slice(i, open.length + i) == open) {
                                    i += open.length;
                                    var pixEq = str.substr(i, 1);
                                    var prefix, postfix, line = lineno;
                                    if (pixEq === "=" || pixEq === "-" || pixEq === "~" || pixEq === "#") {
                                        prefix = "', (" + line + ", ";
                                        postfix = "), '";
                                        ++ i;
                                    } else {
                                        prefix = "');" + line + ";";
                                        postfix = "; buf.push('";
                                    }
                                    var end = str.indexOf(close, i), js = str.substring(i, end), start = i, include = null, n = 0;
                                    if ("-" == js[js.length - 1]) {
                                        js = js.substring(0, js.length - 2);
                                        consumeEOL = true;
                                    }
                                    return _builder_$0.f(
                                        _builder_$0.e(function() {
                                            if (0 == js.trim().indexOf("include")) {
                                                var name = js.trim().slice(7).trim();
                                                if (! filename) 
                                                    return _builder_$0.k(new Error("filename option is required for includes"));
                                                var path = resolveInclude(name, filename);
                                                return _builder_$0.n(readAsync(path, "utf8"), function (_result_$) {
                                                    include = _result_$;
                                                    extname = path.extname(path);
                                                    return _builder_$0.f(
                                                        _builder_$0.e(function() {
                                                            if (extname === "ejs") {
                                                                return _builder_$0.n(exports.parse(include, {
                                                                    "filename": path,
                                                                    "open": open,
                                                                    "close": close
                                                                }), function (_result_$) {
                                                                    include = _result_$;
                                                                    buf.push("' + (function(){" + include + "})() + '");
                                                                    return _builder_$0.h();
                                                                });
                                                            } else {
                                                                buf.push("'" + include + " '");
                                                                return _builder_$0.h();
                                                            }
                                                        }),
                                                        _builder_$0.e(function() {
                                                            js = "";
                                                            return _builder_$0.h();
                                                        })
                                                    );
                                                });
                                            } else {
                                                return _builder_$0.h();
                                            }
                                        }),
                                        _builder_$0.e(function() {
                                            while (~ (n = js.indexOf("\n", n))) 
                                                (n ++, lineno ++);
                                            if (js.substr(0, 1) == ":") 
                                                js = filtered(js);
                                            if (js) {
                                                if (js.lastIndexOf("//") > js.lastIndexOf("\n")) 
                                                    js += "\n";
                                                if (reXmlEq !== undefined && (pixEq === "=" || pixEq === "-" || pixEq === "~" || pixEq === "#")) 
                                                    js = reXmlEq(pixEq, js);
                                                buf.push(prefix, js, postfix);
                                            }
                                            i += end - start + close.length - 1;
                                            return _builder_$0.h();
                                        })
                                    );
                                } else if (str.substr(i, 1) == "\\") {
                                    buf.push("\\\\");
                                    return _builder_$0.h();
                                } else if (str.substr(i, 1) == "'") {
                                    buf.push("\\'");
                                    return _builder_$0.h();
                                } else if (str.substr(i, 1) == "\r") {
                                    buf.push(" ");
                                    return _builder_$0.h();
                                } else if (str.substr(i, 1) == "\n") {
                                    if (consumeEOL) {
                                        consumeEOL = false;
                                    } else {
                                        buf.push("\\n");
                                        lineno ++;
                                    }
                                    return _builder_$0.h();
                                } else {
                                    buf.push(str.substr(i, 1));
                                    return _builder_$0.h();
                                }
                            });
                        })
                    );
                }),
                _builder_$0.e(function() {
                    buf.push("');\nreturn buf.join('');");
                    return _builder_$0.g(buf.join(""));
                })
            );
        })
    );
});

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
