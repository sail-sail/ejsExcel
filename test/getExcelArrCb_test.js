var ejsExcel = require("../ejsExcel");
var fs = require("fs");
//获得Excel模板的buffer对象
var exlBuf = fs.readFileSync("./test.xlsx");

ejsExcel.getExcelArrCb(exlBuf,function(exlJson) {
	  console.log(exlJson);
});