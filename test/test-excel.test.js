var fs = require("fs");
var path=require('path');
var ejsExcel = require("../ejsExcel");

var TEMPLATE_PATH=path.join(__dirname,"./template/test.xlsx");
var OUT_PATH=path.join(__dirname,"./generated/test2.xlsx");
describe('test excel 1', function () {

    it('test test.xlsx',function () {
        //获得Excel模板的buffer对象
        var exlBuf = fs.readFileSync(TEMPLATE_PATH);

        //数据源
        var data = [
            [{ "dpt_des": "开发部", "doc_dt": "2013-09-09", "doc": "a001" }], 
            [
                { "pt": "pt1", "des": "des1", "due_dt": "2013-08-07", "des2": "2013-12-07" }, 
                { "pt": "pt1", "des": "des1", "due_dt": "2013-09-14", "des2": "des21" }
            ]
        ];

        //用数据源(对象)data渲染Excel模板
        return ejsExcel.renderExcel(exlBuf, data)
            .then(function (exlBuf2) {
                fs.writeFileSync(OUT_PATH, exlBuf2);
                console.log("生成test2.xlsx");
            }).catch(function (err) {
                console.error(err);
            });
    });
});