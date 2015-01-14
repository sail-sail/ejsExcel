//把buffer左边的0全部去除,返回新的起点
exports.delLeft0 = function(buf){
	var i;
	for(i=0; i<buf.length; i++) if(buf[i] !== 0x00) break;
	return i;
};
exports.replaceBuf = function(begin,end,buf,buf2){
	if(buf2 === undefined || buf2 === null) {
		buf2 = new Buffer(0);
	}
	var buffer = new Buffer(buf.length-(end-begin)+buf2.length);
	for(var i=0; i<buffer.length; i++) {
		if(i < begin) {
			buffer[i] = buf[i];
		}
		else if(i >= begin && i < begin+buf2.length) {
			buffer[i] = buf2[i-begin];
		}
		else if(i >= begin+buf2.length) {
			buffer[i] = buf[i-buf2.length+(end-begin)];
		}
	}
	return buffer;
};
//var buffer = new Buffer(4);
//buffer[0] = 0x00;
//buffer[1] = 0x01;
//buffer[2] = 0x02;
//buffer[3] = 0x03;
//var buffer2 = new Buffer(1);
//buffer2[0] = 0x2B;
//console.log(exports.replaceBuf(1,3,buffer,buffer2));