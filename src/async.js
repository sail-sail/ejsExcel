
/**
 * convert a generator function `fn` to another function which will call generator.next() in sequence (via promise)
 * @param {Generator} fn 
 */
function _asyncToGenerator(fn) {
    return function () {

        var gen = fn.apply(this, arguments);
        return new Promise(function (resolve, reject) {

            /**
             * @param {String} key "next"||"throw"
             * @param {Object} arg value || err
             */
            function step(key, arg) {
                try {
                    // info=gen.next(arg) or info=gen.throw(arg)
                    var info = gen[key](arg);
                    var value = info.value;
                }
                catch (error) {
                    reject(error); return;
                }
                if (info.done) { return resolve(value); }
                return Promise.resolve(value)
                    .then(
                        function (value) { return step("next", value); },
                        function (err) { return step("throw", err); } // the value is thenable ? 
                    );
            }
            return step("next");
        });
    };
}

/**
 * promisify function . The function receives a callback parameter like `(data)=>{}` 
 * @param {Function} fn 
 * @param {*} thisObj 
 */
function Promise_fromCallback(fn, thisObj) {
    return function () {
        var args= Array.from(arguments);
        if (!thisObj) { thisObj = this; }
        return new Promise(function (resolve, reject) {
            // the new argments is  [...args, callback]
            args.push(function (data) {
                resolve(data);
            });
            if (fn) { fn.apply(thisObj, args); } 
            else { resolve(); }
        });
    };
};
  
/**
 * promisify function . The function receives a callback parameter like `(err,data)=>{}` 
 * @param {Function} fn 
 * @param {*} thisObj 
 */  
function Promise_fromStandard(fn, thisObj) {
    return function () {
        var args = Array.from(arguments);
        if (!thisObj) { thisObj = this; }
        return new Promise(function (resolve, reject) {
            // the new argments is  [...args, callback]
            args.push(function (err, data) {
                if (err) { reject(err); return; }
                resolve(data);
            });
            if (fn) { fn.apply(thisObj, args); } 
            else { resolve(); }
        });
    };
};
  
/**
 * timeout function with promise style
 * @param {Number} time 
 * @return {Promise}
 */
function Promise_sleep(time) {
    return new Promise(function (resolve, reject) {
        if (time > 0) {
        setTimeout(function () { resolve(); }, time);
        } else {
        setImmediate(function () { resolve(); });
        }
    });
};



module.exports={
    _asyncToGenerator,
    Promise_fromCallback,
    Promise_fromStandard,
    Promise_sleep,
};