var fs = require("fs");
var path=require('path');
var ejsExcel = require("../ejsExcel");


var EXCEL_PATH=path.join(__dirname,"./test.xlsx");
describe('test get Excel Attribute', function () {

    it('test test.xlsx',function () {
        //获得Excel模板的buffer对象
        var exlBuf = fs.readFileSync(EXCEL_PATH);

        //getExcelArr 返回Promise对象
        return ejsExcel.getExcelArr(exlBuf)
            .then(function(exlJson) {
                console.log(exlJson);
            }).catch(function(err){
                console.error(err);
            });
    });
});