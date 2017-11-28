

function isType(type) {
    return function (obj) {
        return Object.prototype.toString.call(obj) === "[object " + type + "]";
    };
};

var isObject = isType("Object");
var isString = isType("String");
var isArray = Array.isArray || isType("Array");
var isFunction = isType("Function");



module.exports={
    isType,
    isObject,
    isString,
    isArray,
    isFunction,
};