var ejsExcel = require("../ejsExcel");
var fs = require("fs");
//获得Excel模板的buffer对象
var exlBuf = fs.readFileSync("./test.xlsx");

//数据源
var data = [[{"dpt_des":"开发部","doc_dt":"2013-09-09","doc":"a001"}],[{"pt":"pt1","des":"des1","due_dt":"2013-08-07","des2":"2013-12-07"},{"pt":"pt1","des":"des1","due_dt":"2013-09-14","des2":"des21"}]];

//用数据源(对象)data渲染Excel模板
ejsExcel.renderExcelCb(exlBuf, data, function(err,exlBuf2){
  if(err) {
	  console.error(err);
	  return;
  }
  fs.writeFileSync("./test2.xlsx", exlBuf2);
  console.log("生成test2.xlsx");
});
