#spring-boot-poi

使用 Apache POI 来处理office 文件

端口：8888

运行后可访问 http://localhost:8888

## Excel 文件导出
导出测试Excel文件：http://localhost:8888/excel/download
从对象列表导出学生名单：http://localhost:8888/excel/students

## Excel 文件导入
[进入首页](http://localhost:8888)，选择本地Excel 文件后，点击提交。

## 导入学生名单，解析后返回 json