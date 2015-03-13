var fs = require("fs");
var crc = require("./crc32");
var zlib = require("zlib");
var BufferExt = require("./BufferExt");

//sail,2012-09-26
var Hzip = function(buffer){
	var t = this;
	t.buffer = buffer;
	t.getEntries();
};
Hzip.prototype.getEntries = function() {
	var t = this;
	var buffer = t.buffer;
	var entries = t.entries = [];
	var entries0102 = t.entries0102 = [];
	var tmp = 0;
	while(true) {
		if(tmp > buffer.length) break;
		//0~3分文件头信息标志(\x50\x4b\x03\x04)
		if(buffer[tmp] !== 0x50 || buffer[tmp+1] !== 0x4B) {
			tmp ++;
			continue;
		}
		if(buffer[tmp+2] === 0x03 && buffer[tmp+3] === 0x04) {
			var entry = {};
			tmp = t.getEntry0304(tmp,entry);
			entries.push(entry);
		} else if(buffer[tmp+2] === 0x01 && buffer[tmp+3] === 0x02) {
			var entry = {};
			tmp = t.getEntry0102(tmp,entry);
			entries0102.push(entry);
		}
		else {
			tmp ++;
			continue;
		}
	}
	return t;
};
Hzip.prototype.getEntry0102 = function(tmp,entry){
	var t = this;
	var buffer = t.buffer;
	entry.begin = tmp;
	tmp += 4;
	//4	2	Version made by
	entry.versionMadeBy = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//解压缩所需版本(\x14\x00)  6	2	Version needed to extract (minimum)
	entry.unzipVersion = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//通用比特标志位(置比特0位=加密;置比特1位=使用压缩方式6,并使用8k变化目录,否则使用4k变化目录;置比特2位=使用压缩方式6,并使用3个ShannonFano树对变化目录输出编码,否则使用2个ShannonFano树对变化目录输出编码,其它比特位未用)  
	//(\x00\x00) 8	2	General purpose bit flag
	entry.generalPurpose = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//压缩方式(0=不压缩,1=缩小,2=以压缩因素1缩小,3=以压缩因素2缩小,4=以压缩因素3缩小,5=以压缩因素4缩小,6=自展)
	//(\x08\x00)10	2	Compression method
	entry.compressionWay = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//文件最后修改时间12	2	File last modification time
	entry.lastModifyTime = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//文件最后修改日期14	2	File last modification date
	entry.lastModifyDate = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//32位校验码16	4	CRC-32
	entry.crc32 = buffer.slice(tmp,tmp+4);
	tmp += 4;
	//压缩文件大小20	4	Compressed size
	entry.cfileSize = buffer.slice(tmp,tmp+4).readInt32LE(0);
	tmp += 4;
	//未压缩文件大小24	4	Uncompressed size
	entry.fileSize = buffer.slice(tmp,tmp+4).readInt32LE(0);
	tmp += 4;
	//文件名长28	2	File name length (n)
	var n = entry.fileNameSize = buffer.slice(tmp,tmp+2).readInt16LE(0);
	tmp += 2;
	//扩展段长30	2	Extra field length (m)
	var m = entry.extraFieldSize = buffer.slice(tmp,tmp+2).readInt16LE(0);
	tmp += 2;
	//32	2	File comment length (k)
	var k = entry.fileCommentSize = buffer.slice(tmp,tmp+2).readInt16LE(0);
	tmp += 2;
	//34	2	Disk number where file starts
	var k = entry.diskNumStarts = buffer.slice(tmp,tmp+2).readInt16LE(0);
	tmp += 2;
	//36	2	Internal file attributes
	entry.internalFileAttrs = buffer.slice(tmp,tmp+2).readInt16LE(0);
	tmp += 2;
	//38	4	External file attributes
	entry.externalFileAttrs = buffer.slice(tmp,tmp+4).readInt32LE(0);
	tmp += 4;
	//42	4	Relative offset of local file header. This is the number of bytes between the start of the first disk on which the file occurs, and the start of the local file header. This allows software reading the central directory to locate the position of the file inside the ZIP file.
	entry.offsetOfHeader = buffer.slice(tmp,tmp+4).readInt32LE(0);
	tmp += 4;
	//46	n	File name
	entry.fileName = buffer.slice(tmp,tmp+n).toString();
	tmp += n;
	//46+n	m	Extra field
	entry.extraField = buffer.slice(tmp,tmp+m);
	tmp += m;
	//46+n+m	k	File comment
	entry.fileComment = buffer.slice(tmp,tmp+k).toString();
	tmp += k;
	entry.end = tmp;
	return tmp;
};
Hzip.prototype.getEntry0304 = function(tmp,entry){
	var t = this;
	var buffer = t.buffer;
	entry.begin = tmp;
	tmp += 4;
	//4~6解压缩所需版本(\x14\x00) version needed to extract       2 bytes
	entry.unzipVersion = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//6~8通用比特标志位(置比特0位=加密;置比特1位=使用压缩方式6,并使用8k变化目录,否则使用4k变化目录;置比特2位=使用压缩方式6,并使用3个ShannonFano树对变化目录输出编码,否则使用2个ShannonFano树对变化目录输出编码,其它比特位未用)  
	//(\x00\x00) general purpose bit flag        2 bytes
	entry.generalPurpose = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//8~10压缩方式(0=不压缩,1=缩小,2=以压缩因素1缩小,3=以压缩因素2缩小,4=以压缩因素3缩小,5=以压缩因素4缩小,6=自展)
	//(\x08\x00)
	entry.compressionWay = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//10~12文件最后修改时间
	entry.lastModifyTime = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//12~14文件最后修改日期
	entry.lastModifyDate = buffer.slice(tmp,tmp+2);
	tmp += 2;
	//14~18,32位校验码
	entry.crc32 = buffer.slice(tmp,tmp+4);
	tmp += 4;
	//18~22压缩文件大小
	entry.cfileSize = buffer.slice(tmp,tmp+4).readInt32LE(0);
	tmp += 4;
	//22~26未压缩文件大小
	entry.fileSize = buffer.slice(tmp,tmp+4).readInt32LE(0);
	tmp += 4;
	//26~28文件名长
	entry.fileNameSize = buffer.slice(tmp,tmp+2).readInt16LE(0);
	tmp += 2;
	//28~30扩展段长
	entry.extraFieldSize = buffer.slice(tmp,tmp+2).readInt16LE(0);
	tmp += 2;
	//31~30+entry.fileNameSize文件名
	entry.fileName = buffer.slice(tmp,tmp+entry.fileNameSize).toString();
	tmp += entry.fileNameSize;
	//30+fileNameSize~30+fileNameSize+extraFieldSize扩展段
	var b20a2 = buffer.slice(tmp,tmp+2);
	var extraFieldSize = entry.extraFieldSize;
	if(extraFieldSize > 0) {
		entry.extraField = buffer.slice(tmp,tmp+extraFieldSize);
	} else extraFieldSize = 0;
	if(b20a2[0] === 0x20 && b20a2[1] === 0xA2) {
		tmp += 2;
		var tg = buffer.slice(tmp,tmp+2).readInt16LE(0);
		tmp += 2;
		tmp += tg;
	} else tmp += extraFieldSize;
	entry.cfile = buffer.slice(tmp,tmp+entry.cfileSize);
	tmp += entry.cfileSize;
	entry.end = tmp;
	return tmp;
};
Hzip.prototype.getEntry = function(fileName,type){
	type = type || "";
	var t = this;
	var entries = t["entries"+type];
	var entry = undefined;
	for(var i=0; i<entries.length; i++) {
		if(fileName === entries[i].fileName) {
			entry = entries[i];
			break;
		}
	}
	return entry;
};
//更新或者增加entry之后,要重新设置0102和0506
Hzip.prototype.zip = function(fileComment,comment) {
	fileComment = fileComment || "";
	comment = comment || "";
	var t = this;
	t.getEntries();
	var entries0102 = [];
	for(var i=0; i<t.entries.length; i++) {
		var entry = t.entries[i];
		var n = entry.fileNameSize;
		//Extra field length (m)
		var m = 0;
		var k = fileComment.length;
		var entry0102 = new Buffer(46+n+m+k);
		entries0102.push(entry0102);
		//Central directory file header signature = 0x02014b50
		entry0102[0] = 0x50;
		entry0102[1] = 0x4B;
		entry0102[2] = 0x01;
		entry0102[3] = 0x02;
		//Version made by
		entry0102[4] = 0x2D;
		entry0102[5] = 0x00;
		//Version needed to extract (minimum) 14 00
		entry0102[6] = entry.unzipVersion[0];
		entry0102[7] = entry.unzipVersion[1];
		//8	2	General purpose bit flag 06 00
		entry0102[8] = entry.generalPurpose[0];
		entry0102[9] = entry.generalPurpose[1];
		//10	2	Compression method 08 00
		entry0102[10] = entry.compressionWay[0];
		entry0102[11] = entry.compressionWay[1];
		//12	2	File last modification time 00 00
		entry0102[12] = entry.lastModifyDate[0];
		entry0102[13] = entry.lastModifyDate[1];
		//14	2	File last modification date 21 00
		entry0102[14] = entry.lastModifyDate[0];
		entry0102[15] = entry.lastModifyDate[1];
		//16	4	CRC-32
		entry0102[16] = entry.crc32[0];
		entry0102[17] = entry.crc32[1];
		entry0102[18] = entry.crc32[2];
		entry0102[19] = entry.crc32[3];
		//20	4	Compressed size cfileSize
		entry0102.writeInt32LE(entry.cfileSize,20);
		//24	4	Uncompressed size fileSize
		entry0102.writeInt32LE(entry.fileSize,24);
		//28	2	File name length (n)
		entry0102.writeInt16LE(n,28);
		//30	2	Extra field length (m)
		entry0102.writeInt16LE(m,30);
		//32	2	File comment length (k)
		entry0102.writeInt16LE(k,32);
		//34	2	Disk number where file starts
		entry0102.writeInt16LE(0,34);
		//36	2	Internal file attributes
		entry0102[36] = 0x00;
		entry0102[37] = 0x00;
		//38	4	External file attributes
		entry0102[38] = 0x00;
		entry0102[39] = 0x00;
		entry0102[40] = 0x00;
		entry0102[41] = 0x00;
		//42	4	Relative offset of local file header. This is the number of bytes between the start of the first disk on which the file occurs, and the start of the local file header. This allows software reading the central directory to locate the position of the file inside the ZIP file.
		entry0102.writeInt32LE(entry.begin,42);
		//46	n	File name
		var fileNameBuf = new Buffer(entry.fileName);
		for(var j=0; j<n; j++) {
			entry0102[46+j] = fileNameBuf[j];
		}
		//46+n	m	Extra field
		//46+n+m	k	File comment
		var fileCommentBuf = new Buffer(fileComment);
		for(var j=0; j<k; j++) {
			entry0102[46+n+m+j] = fileCommentBuf[j];
		}
	}
	//0506 After all the central directory entries comes the end of central directory record, which marks the end of the ZIP file:
	var entry0506 = new Buffer(22+comment.length);
	// 0	4	End of central directory signature = 0x06054b50
	entry0506[0] = 0x50;
	entry0506[1] = 0x4B;
	entry0506[2] = 0x05;
	entry0506[3] = 0x06;
	// 4	2	Number of this disk
	entry0506[4] = 0x00;
	entry0506[5] = 0x00;
	//6	2	Disk where central directory starts
	entry0506[6] = 0x00;
	entry0506[7] = 0x00;
	//8	2	Number of central directory records on this disk,entries0102.length
	entry0506.writeInt16LE(entries0102.length,8);
	//10	2	Total number of central directory records
	entry0506.writeInt16LE(entries0102.length,10);
	//12	4	Size of central directory (bytes)
	var entries0102Size = 0;
	for(var i=0; i<entries0102.length; i++) {
		entries0102Size += entries0102[i].length;
	}
	entry0506.writeInt32LE(entries0102Size,12);
	//16	4	Offset of start of central directory, relative to start of archive
	var entriesSize = 0;
	for(var i=0; i<t.entries.length; i++) {
		entriesSize = entriesSize + t.entries[i].end - t.entries[i].begin;
	}
	entry0506.writeInt32LE(entriesSize,16);
	//20	2	Comment length (n)
	entry0506.writeInt16LE(comment.length,20);
	//22	n	Comment
	var commentBuf = new Buffer(comment);
	for(var i=0; i<commentBuf.length; i++) {
		entry0506[22+i] = commentBuf[i];
	}
	
	//重新组合buffer
	var lenbuf1 = 0;
	lenbuf1 += entriesSize;
	for(var i=0; i<entries0102.length; i++) {
		var entry0102 = entries0102[i];
		lenbuf1 += entry0102.length;
	}
	lenbuf1 += entry0506.length;
	var buffer1 = new Buffer(lenbuf1);
	var targetStart = 0;
	for(var i=0; i<t.entries.length; i++) {
		t.buffer.copy(buffer1,targetStart,t.entries[i].begin,t.entries[i].end);
		targetStart = targetStart + t.entries[i].end - t.entries[i].begin;
	}
	for(var i=0; i<entries0102.length; i++) {
		var entry0102 = entries0102[i];
		entry0102.copy(buffer1,targetStart);
		targetStart += entry0102.length;
	}
	entry0506.copy(buffer1,targetStart);
	t.buffer = buffer1;
	t.getEntries();
};
Hzip.prototype.removeEntry = function(fileName){
	var t = this;
	var entry = t.getEntry(fileName);
	if(entry !== undefined && entry !== null) {
		t.buffer = BufferExt.replaceBuf(entry.begin,entry.end,t.buffer);
		t.zip();
	}
};
//更新zip压缩包里面的文件filename
Hzip.prototype.updateEntry = function(fileName,fileBuf,isDefRaw,callback){
	var t = this;
	if(callback === undefined && typeof isDefRaw === "function") callback = isDefRaw;
	if(fileBuf === undefined || fileBuf === null) {
		t.removeEntry(fileName);
		if(callback) callback(null,t.buffer);
		return;
	}
	if(!Buffer.isBuffer(fileBuf)) {
		fileBuf = new Buffer(fileBuf);
	}
	t.toEntryBuf(fileName,fileBuf,isDefRaw,function(err,buf){
		var entry = t.getEntry(fileName);
		var begin = 0;
		var end = 0;
		if(entry) {
			begin = entry.begin;
			end = entry.end;
		}
		t.buffer = BufferExt.replaceBuf(begin,end,t.buffer,buf);
		t.zip();
		if(callback) callback(err,t.buffer);
	});
};
//isDefRaw 是否压缩,默认压缩,sail 2014-01-13
Hzip.prototype.toEntryBuf = function(fileName,fileBuf,isDefRaw,callback){
	var c32Num = crc.crc32(fileBuf);
	var fileLength = fileBuf.length;
	if(!Buffer.isBuffer(fileName)) fileName = new Buffer(fileName);
	var tmpFn = function(err,cfile){
		var eb = new Buffer(30+fileName.length+cfile.length);
		eb[0] = 0x50;
		eb[1] = 0x4B;
		eb[2] = 0x03;
		eb[3] = 0x04;
		if(isDefRaw !== false) {
			//4~5解压缩所需版本(\x14\x00)
			eb[4] = 0x14;
			eb[5] = 0x00;
			//6~7通用比特标志位(置比特0位=加密;置比特1位=使用压缩方式6,并使用8k变化目录,否则使用4k变化目录;置比特2位=使用压缩方式6,并使用3个ShannonFano树对变化目录输出编码,否则使用2个ShannonFano树对变化目录输出编码,其它比特位未用)  
			//(\x00\x00)
			eb[6] = 0x06;
			eb[7] = 0x00;
			//8~9解压缩所需版本(\x08\x00)
			eb[8] = 0x08;
			eb[9] = 0x00;
		} else {
			eb[4] = 0x0A;
			eb[5] = 0x00;
			eb[6] = 0x00;
			eb[7] = 0x00;
			eb[8] = 0x00;
			eb[9] = 0x00;
		}
		//10~11文件最后修改时间
		eb[10] = 0x00;
		eb[11] = 0x00;
		//12~13文件最后修改日期
		eb[12] = 0x21;
		eb[13] = 0x00;
		//14~17,32位校验码
		eb.writeInt32LE(c32Num,14);
		//18~21压缩文件大小
		eb.writeInt32LE(cfile.length,18);
		//22~25未压缩文件大小
		eb.writeInt32LE(fileLength,22);
		//26~27File name length
		eb.writeInt16LE(fileName.length,26);
		//28~29Extra field length
		eb.writeInt16LE(0,28);
		//30 file name
		for(var i=0; i<fileName.length; i++) {
			eb[i+30] = fileName[i];
		}
		//Extra field
		//cfile
		for(var i=0; i<cfile.length; i++) {
			eb[i+30+fileName.length] = cfile[i];
		}
		callback(err,eb);
	};
	if(isDefRaw === false) {
		tmpFn(null,fileBuf);
	} else {
		zlib.deflateRaw(fileBuf,tmpFn);
	}
};
module.exports = Hzip;
