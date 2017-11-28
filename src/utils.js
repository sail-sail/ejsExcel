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



module.exports={
    charToNum,
    charPlus,
};