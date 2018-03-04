#spring-boot-poi

使用 Apache POI 来处理office 文件

端口：8888

运行后可访问 http://localhost:8888

## Excel 文件导出
创建并下载测试Excel文件：http://localhost:8888/excel/create-xls
从对象列表导出学生名单：http://localhost:8888/excel/students

## Excel 文件导入
[进入首页](http://localhost:8888)，选择本地Excel 文件后，点击提交。

## 导入学生名单，解析后返回 json
[进入首页](http://localhost:8888) ,点击“下载学生名单”，获得学生名单的Excel文件，在表单中选择下载的名单文件，点击提交，将返回json格式数据。

## 创建 xlsx 文件
[进入首页](http://localhost:8888) ,点击“创建xlsx文件”。

## 解析 xlsx 文件
[进入首页](http://localhost:8888) ,在表单中选择xlsx文件，点击提交，将返回json格式数据。

## 设置单元格对齐方式