ejsExcel
--------
> nodejs excel template engine. node export excel, ejsExcel


### How to use?
```bash
npm install ejsexcel
```
   
### How to test?
- 执行 test/test.bat
    ```bash
    test/test.bat
    ```
- test/test.xlsx 为完整示例 demo

- e.g
   ```js
    const ejsexcel = require("ejsexcel");
    const fs = require("fs");
    const util = require("util");
    const readFileAsync = util.promisify(fs.readFile);
    const writeFileAsync = util.promisify(fs.writeFile);
    
   (async function() {
     //获得Excel模板的buffer对象
     const exlBuf = await readFileAsync("./test.xlsx");
     //数据源
     const data = [];
     //用数据源(对象)data渲染Excel模板
     //cachePath 为编译缓存路径, 对于模板文件比较大的情况, 可显著提高运行效率, 绝对路径, 若不设置, 则无缓存
     const exlBuf2 = await ejsexcel.renderExcel(exlBuf, data, { cachePath: __dirname+"/cache/" });
     await writeFileAsync("./test2.xlsx", exlBuf2);
     console.log("生成test2.xlsx");
   })();
   ```

## Syntax

| Syntax                | Description                               |
|-----------------------|-------------------------------------------|
| _data_                | _data_ 为内置对象, 数据源                   |
| <%forRow              | 循环一行                                  |
| <%#                   | 输出动态公式                               |
| <%~                   | 输出数字类型格式                           |
| <%=                   | 输出字符串                                |
| <%forCell             | 循环单元格                                |
| <%forRBegin           | 循环多行                                  |
| <%forCBegin           | 循环多个单元格                             |
| <%_hideSheet_()%>     | 隐藏所在工作表                             |
| <%_showSheet_()%>     | 显示所在工作表                             |
| <%_deleteSheet_()%>   | 删除所在工作表                             |
| <%   %>               | 内部可执行 任意 javascript,可以用 <%console.log(_data_)%> 打印临时变量到控制台,进行调试 |

## Author
+ Author: Sail, 辐毂  
    - QQ: 151263555  
    - QQ群: 470988427  
    - email: 151263555@qq.com 

## Templates
> 做一个这样的模版
![模板](http://dn-cnode.qbox.me/Frs_RuLXJxYQgYoIUhGJJ1zspCJE)

## Result
> ## 加数据渲染之后,导出结果
![导出结果](http://dn-cnode.qbox.me/FnRDa5Zyjg-dI7ykCNR0T8SorWyC)


## 捐赠鼓励支持此项目,支付宝扫描:
![捐赠鼓励支持此项目](http://dn-cnode.qbox.me/FucPKV4XWewhakoqTSngU3AsaP0Z)

## 项目贡献人列表
- @Hello World  ¥50
- @德爾文  ¥50
- @Explore®  ¥50
- @向左转  ¥50
- @吴燕飞  ¥50
- @strive-ming  ¥10
- @MR.P  ¥16.66
- @不求来生  ¥6.60
- @羊刀  ¥6.66
- @Leo  ¥8.88
- @不摇碧莲  ¥6.66
- @天涯海角  ¥10
- @刚刚好  ¥5
- @Aptenodytes forsteri  ¥15
- @HHB  ¥30
- @Original  ¥88.88
- @半日  ¥10
- @狗狗踏青去了  ¥50
- @zzzzzzzz  ¥50
- @zzzzzzzz  ¥30
- @九尾之狐  ¥66.66
