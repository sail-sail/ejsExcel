var ejsExcel = require("../ejsExcel");
var fs = require("fs");

//获得Excel模板的buffer对象
var exlBuf = fs.readFileSync("./order_tmpl.xlsx");

//数据源
var data = {"name":"test",open:"<%",close:"%>"};

//用数据源(对象)data渲染Excel模板
ejsExcel.renderExcelCb(exlBuf, data, function(exlBuf2){
  fs.writeFileSync("./order2.xlsx", exlBuf2);
  console.log("生成order2.xlsx");
});