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
    Promise_fromCallback,
    Promise_fromStandard,
    Promise_sleep,
};