const {isString}=require('./is-type');

/**
 * Convert str to number 
 * @param {String} str 
 */
function charToNum (str) {
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
 * add delta to char
 * @param {String} str 
 * @param {Number} num 
 */
function charPlus(str, num) {
    var ch, i, strNum, temp;
    strNum = charToNum(str);
    strNum += num;
    if (strNum <= 0) {
        return "A";
    }
    temp = "";
    ch = "";
    while (strNum >= 1) {
        i = strNum % 26;
        if (i !== 0) {
            ch = String.fromCharCode(65 + i - 1);
            temp = ch + temp;
        } else {
            ch = "Z";
            temp = ch + temp;
            strNum--;
        }
        strNum = Math.floor(strNum / 26);
    }
    return temp;
};


/**
 * escape string to xml
 * @param {String} str 
 */
function str2Xml(str) {
    if (!isString(str)) {
        return str;
    }
    return Array.from(str,(s,k)=>{
        if (s === "&") {
            s = "&amp;";
        } else if (s === "<") {
            s = "&lt;";
        } else if (s === ">") {
            s = "&gt;";
        } else if (s === "\"") {
            s = "&quot;";
        } else if (s === "'") {
            s = "&apos;";
        }
        return s;
    }).join('');
};


/**
 * 
 * @param {String} tt main string 
 * @param {*} what substring (pattern) to be replaced
 * @param {*} replacement  replaced with
 */
function replaceLast(tt, what, replacement) {
    if(!tt) return;
    let mthArr = tt.match(what);
    let num = 0;
    const result= tt.replace(what, function (s) {
        num++;
        if (num === mthArr.length) {
            return replacement;
        }
        return s;
    });
    return result;
};

module.exports={
    charToNum,
    charPlus,
    str2Xml,
    replaceLast,
};