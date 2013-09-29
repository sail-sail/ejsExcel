hzip
====

```javascript
var fs = require("fs");
var zlib = require("zlib");
var Hzip = require("../hzip");
var hzip = new Hzip(fs.readFileSync("./test.zip"));
//替换或增加文件
hzip.updateEntry("testDir/test.txt",fs.readFileSync("./test.txt"),function(err,buffer){
	if(err) console.log(err);
	if(fs.existsSync("./test2.zip") === true) fs.unlinkSync("./test2.zip");
	fs.writeFileSync("./test2.zip",buffer);
	//解压文件
	var entry = hzip.getEntry("testDir/test.txt");
	zlib.inflateRaw(entry.cfile,function(err,buf){
		console.log(buf.toString());
	});
});
```
由于Excel2007以上的版本其实都是zip文件,里面就是xml文档
那么,你懂的,它可以用来导入导出Excel2007或者以上的版本
实际项目中已经在用了
