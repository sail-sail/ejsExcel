### nodejs excel template engine. node export excel, ejsExcel
   
**npm install ejsexcel**

```
执行 test/test.bat  进行测试
test/test.xlsx为完整示例
```

```
       	   <%forRow 循环一行
		    	<%# 输出动态公式
		    	<%~ 输出数字类型格式
		    	<%= 输出字符串
	  	  <%forCell 循环单元格
		<%forRBegin 循环多行
		<%forCBegin 循环多个单元格
  <%_hideSheet_()%> 隐藏所在工作表
<%_deleteSheet_()%> 删除所在工作表

	此外: <% %>内部可执行 任意 javascript, 注意是 **任意**
          _data_为内置对象, 数据源
	比如: 可以用<%console.log(_data_)%>
         打印临时变量到控制台,进行调试
```
====  
```
auto: Sail  
QQ: 151263555  
QQ群: 470988427  
email: 151263555@qq.com 
```
  
**模板:**

![模板](http://dn-cnode.qbox.me/Frs_RuLXJxYQgYoIUhGJJ1zspCJE)

**加数据渲染之后,导出结果:**

![导出结果](http://dn-cnode.qbox.me/FnRDa5Zyjg-dI7ykCNR0T8SorWyC)

捐赠鼓励支持此项目:

![捐赠鼓励支持此项目](http://dn-cnode.qbox.me/FucPKV4XWewhakoqTSngU3AsaP0Z)

**项目贡献人列表:**

*   @Hello World~&nbsp;&nbsp;&nbsp;&nbsp;50元
*   @德爾文&nbsp;&nbsp;&nbsp;&nbsp;50元