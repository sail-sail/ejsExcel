var ejsExcel = require("../ejsExcel");
var fs = require("fs");
//获得Excel模板的buffer对象
var exlBuf = fs.readFileSync("./test.xlsx");

//数据源
var data = [[{"dpt_des":"开发部","doc_dt":"2013-09-09","doc":"a001"}],[{"pt":"pt1","des":"des1","due_dt":"due_dt1","des2":"des21"},{"pt":"pt1","des":"des1","due_dt":"due_dt1","des2":"des21"}]];

//用数据源(对象)data渲染Excel模板
ejsExcel.renderExcelCb(exlBuf, data, function(exlBuf2){
  fs.writeFileSync("./test2.xlsx", exlBuf2);
});
