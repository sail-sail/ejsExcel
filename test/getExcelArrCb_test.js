var fs = require("fs");
var path=require('path');
var ejsExcel = require("../ejsExcel");


var TEMPLATE_PATH=path.join(__dirname,"./template/test.xlsx");
var OUT_PATH=path.join(__dirname,"./generated/test2.xlsx");
describe('test get Excel Attribute', function () {

    it('test test.xlsx',function () {
        //获得Excel模板的buffer对象
        var exlBuf = fs.readFileSync(TEMPLATE_PATH);

        //getExcelArr 返回Promise对象
        return ejsExcel.getExcelArr(exlBuf)
            .then(function(exlJson) {
                console.log(exlJson);
            }).catch(function(err){
                console.error(err);
            });
    });
});