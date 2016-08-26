var ejsExcel = require("../ejsExcel");
var fs = require("fs");
//获得Excel模板的buffer对象
var exlBuf = fs.readFileSync("./test.xlsx");

//getExcelArr 返回Promise对象
ejsExcel.getExcelArr(exlBuf).then(function(exlJson) {
	console.log(exlJson);
}).catch(function(err){
	console.error(err);
});