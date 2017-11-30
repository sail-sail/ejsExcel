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


## 捐赠鼓励支持此项目
![捐赠鼓励支持此项目](http://dn-cnode.qbox.me/FucPKV4XWewhakoqTSngU3AsaP0Z)

## 项目贡献人列表
- @Hello World  ¥50
- @德爾文  ¥50
- @Explore®  ¥50
- @向左转  ¥50