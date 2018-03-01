package xyz.wqf31415.controller;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import xyz.wqf31415.entity.Student;

import javax.servlet.ServletInputStream;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Administrator on 2018/2/26.
 *
 * @author WeiQuanfu
 */
@RestController
@RequestMapping("/excel")
public class ExcelController {

    /**
     * 下载excel表格
     *
     * @param response
     */
    @RequestMapping("/download")
    public void download(HttpServletResponse response){
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet0");
        HSSFRow row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("姓名");
        row0.createCell(1).setCellValue("年龄");
        row0.createCell(2).setCellValue("性别");
        row0.createCell(3).setCellValue("电话");

        HSSFRow row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("张三");
        row1.createCell(1).setCellValue(18);
        row1.createCell(2).setCellValue("男");
        row1.createCell(3).setCellValue("87878787");

        String excelFileName = "details.xls";
        try {
            ServletOutputStream out = response.getOutputStream();
            response.reset();
            response.setHeader("Content-disposition", "attachment; filename=" + excelFileName);
            response.setContentType("application/msexcel");
            workbook.write(out);
            out.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 将实体List导出为Excel
     *
     * @param response
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     * @throws IOException
     */
    @RequestMapping("/students")
    public void downStudentExcel(HttpServletResponse response) throws NoSuchFieldException, IllegalAccessException, IOException {
        // 要导出的数据
        Student s1 = new Student("张三",18,"男");
        Student s2 = new Student("李四",19,"女");
        List<Student> studentList = new ArrayList<>();
        studentList.add(s1);
        studentList.add(s2);

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("学生名单");
        int rowNumber = 0;
        HSSFRow row0 = sheet.createRow(rowNumber++);

        Class studentClass = Student.class;
        Field[] fields = studentClass.getDeclaredFields();
        // 写入首行标题
        for (int i = 0; i<fields.length;i++){
            System.out.println("名称:"+fields[i].getName());
            row0.createCell(i).setCellValue(fields[i].getName());
        }
        // 写入内容
        for (Student stu : studentList){
            HSSFRow row = sheet.createRow(rowNumber++);
            for (int i = 0; i<fields.length;i++){
                Field f = stu.getClass().getDeclaredField(fields[i].getName());
                f.setAccessible(true);
                row.createCell(i).setCellValue(f.get(stu).toString());
            }
        }

        // 导入excel
        String excelFileName = "student.xls";
        ServletOutputStream out = response.getOutputStream();
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=" + excelFileName);
        response.setContentType("application/msexcel");
        workbook.write(out);
        out.close();
    }

    /**
     * 上传 Excel 文件，程序读取内容后以字符串形式返回
     *
     * @param file
     * @return
     * @throws IOException
     */
    @RequestMapping("/upload")
    public String uploadExcelFile(MultipartFile file) throws IOException {
        String result = "";
        InputStream stream = file.getInputStream();
        HSSFWorkbook workbook = new HSSFWorkbook(stream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        // 获取最后一行的索引
        int rowNum = sheet.getLastRowNum();
        for (int i = 0; i <= rowNum; i++){
            HSSFRow row = sheet.getRow(i);
            // 获取单元格数量，即最后一个单元格索引+1
            int cellNum = row.getLastCellNum();
            for (int j = 0; j < cellNum; j++){
                HSSFCell cell = row.getCell(j);
                if (cell != null) {
                    result += cell.getStringCellValue() + "\t";
                }
            }
            result += "\n";
        }
        System.out.println(result);
        return result;
    }
}
