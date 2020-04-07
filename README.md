# sql_go

go批处理excel合并内容。excel数据转sql。

现阶段通过 垃圾代码 实现将excel中一列拼接成sql批量插入、生产等语句。代码逻辑简单，实用。

//TODO待本人go技术精湛后将重构并丰富该工具

编译如下：
```
#加载rsrc，用于预处理manifest文建
go get github.com/akavel/rsrc
rsrc -manifest test.manifest -o main.syso
#编译打包
go build
go build -ldflags="-H windowsgui"
```
打包后产生sql_go.exe

涉及技术

go GUI、excel解析、字符串拼接、类型转换

使用方法请查看release
