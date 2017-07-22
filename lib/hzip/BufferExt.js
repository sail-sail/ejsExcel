exports.replaceBuf = function(begin,end,buf,buf2){
	if(buf2 === undefined || buf2 === null) {
		buf2 = Buffer.allocUnsafe(0);
	}
	return Buffer.concat([buf.slice(0,begin),buf2,buf.slice(end)]);
};
//var buffer = new Buffer(4);
//buffer[0] = 0x00;
//buffer[1] = 0x01;
//buffer[2] = 0x02;
//buffer[3] = 0x03;
//var buffer2 = new Buffer(1);
//buffer2[0] = 0x2B;
//console.log(exports.replaceBuf(1,3,buffer,buffer2));