
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


module.exports={
    charToNum,
};