var expat = require('../../node-expat');
var fs = require('fs');

// This object will hold the final result.
var obj = {}; 
var currentObject = {};
var ancestors = [];
var currentElementName = null;

var options = {}; //configuration options
function startElement(name, attrs) {
    currentElementName = name;
    if(options.coerce) {
        // Looping here in stead of making coerce generic as object walk is unnecessary
        Object.keys(attrs).forEach(function(key) {
            attrs[key] = coerce(attrs[key]);
        });
    }

    if (! (name in currentObject)) {
        currentObject[name] = attrs;
    } else if (! (currentObject[name] instanceof Array)) {
        // Put the existing object in an array.
        var newArray = [currentObject[name]];
        // Add the new object to the array.
        newArray.push(attrs);
        // Point to the new array.
        currentObject[name] = newArray;
    } else {
        // An array already exists, push the attributes on to it.
        currentObject[name].push(attrs);
    }

    // Store the current (old) parent.
    ancestors.push(currentObject);

    // We are now working with this object, so it becomes the current parent.
    if (currentObject[name] instanceof Array) {
        // If it is an array, get the last element of the array.
        currentObject = currentObject[name][currentObject[name].length - 1];
    } else {
        // Otherwise, use the object itself.
        currentObject = currentObject[name];
    }
}

function text(data) {
    //console.log('->' + data + '<-');
    /*if (!data.trim().length) {
        return;
    }*/
    
    if (options.trim) {
        data = data.trim();
    }

    if (options.sanitize) {
    	if(options.sanitizeFn !== undefined) {
    		data = options.sanitizeFn(data);
    	} else {
    		data = sanitize(data);
    	}
    }

    currentObject['$t'] = coerce((currentObject['$t'] || '') + data);
}

function endElement(name) {
    if (currentElementName !== name) {
        delete currentObject['$t'];
    }
    // This should check to make sure that the name we're ending 
    // matches the name we started on.
    var ancestor = ancestors.pop();
    if (!options.reversible) {
        if ((Object.keys(currentObject).length == 1) && ('$t' in currentObject)) {
            if (ancestor[name] instanceof Array) {
                ancestor[name].push(ancestor[name].pop()['$t']);
            } else {
                ancestor[name] = currentObject['$t'];
            }
        }
    }

    currentObject = ancestor;
}

function coerce(value) {
    if (!options.coerce) {
        return value;
    }
    
    var num = Number(value);
    if (!isNaN(num)) {
        return num;
    }

    var _value = value.toLowerCase();
    
    if (_value == 'true' || _value == 'yes') {
        return true;
    }

    if (_value == 'false' || _value == 'no') {
        return false;
    }

    return value;
}


/**
 * Simple sanitization. It is not intended to sanitize 
 * malicious element values.
 *
 * character | escaped
 *      <       &lt;
 *      >       &gt;
 *      (       &#40;
 *      )       &#41;
 *      #       &#35;
 *      &       &amp;
 *      "       &quot;
 *      '       &apos;
 */
//var chars =  {  '&': '&amp;',
//		        '<': '&lt;',
//                '>': '&gt;',
//                '(': '&#40;',
//                ')': '&#41;',
//                '#': '&#35;',
//                '"': '&quot;',
//                "'": '&apos;' };

//sail 2013-03-27
//chars =  {'&lt;'  :'<',
//         '&gt;'  :'>',
//         '&#40;' :'(',
//         '&#41;' :')',
//         '&#35;' :'#',
//         '&amp;' :'&',
//         '&quot;':'"',
//         '&apos;':"'" };
//sail 2013-03-27 end

function sanitize(value) {
    if (typeof value !== 'string') {
        return value;
    }
    
    value = value.replace(new RegExp("&","gm"),"&amp;")
    .replace(new RegExp("<","gm"),"&lt;")
    .replace(new RegExp(">","gm"),"&gt;");

    return value;
}

/**
 * Parses xml to json using node-expat.
 * @param {String|Buffer} xml The xml to be parsed to json.
 * @param {Object} _options An object with options provided by the user.
 * The available options are:
 *  - object: If true, the parser returns a Javascript object instead of
 *            a JSON string.
 *  - reversible: If true, the parser generates a reversible JSON, mainly
 *                characterized by the presence of the property $t.
 *  - sanitize_values: If true, the parser escapes any element value in the xml
 * that has any of the following characters: <, >, (, ), #, #, &, ", '.
 * 
 * @return {String|Object} A String or an Object with the JSON representation
 * of the XML.
 */
module.exports = function(xml, _options) {
	if(_options === undefined) _options = {};
	//sail 2013-04-21
	if(Buffer.isBuffer(xml)) xml = xml.toString();
    var parser = new expat.Parser('UTF-8');

    startElement = _options.startElement || startElement;
    endElement = _options.endElement || endElement;
    parser.on('startElement', startElement);
    parser.on('text', text);
    parser.on('endElement', endElement);

    obj = currentObject = {};
    ancestors = [];
    currentElementName = null;

//    options = {
//        object: false,
//        reversible: false,
//        coerce: true,
//        sanitize: true,
//        trim: true
//    };
    options = {
    	object:true,reversible:true,coerce:false,trim:false,sanitize:false
    };

    for (var opt in _options) {
        options[opt] = _options[opt];
    }

    parser.parse(xml);
//    if (!parser.parse(xml)) {
//        throw new Error('There are errors in your xml file: ' + parser.getError());
//    }

    if (options.object) {
        return obj;
    }

    var json = JSON.stringify(obj);

    //See: http://timelessrepo.com/json-isnt-a-javascript-subset
    json = json.replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
    
    return json;
};
module.exports.startElement = startElement;
module.exports.endElement = endElement;

