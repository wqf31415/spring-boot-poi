package xyz.wqf31415.controller.util;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;
import xyz.wqf31415.entity.Student;

import java.io.File;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import static org.junit.Assert.*;

/**
 * Created by Administrator on 2019/7/11.
 *
 * @author WeiQuanfu
 */
public class ExcelExportUtilTest {

    private final static String exportFilePath = "D:\\test.xlsx";

    @Test
    public void generateWorkbook() throws IOException {
        SXSSFWorkbook workbook = ExcelExportUtil.whiteDate(generateStudentList(),Student.class);


        // 写入文件
        File file = new File(exportFilePath);
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
    }

    private List<Student> generateStudentList(){
        int studentCount = 20;
        Random random = new Random();
        List<Student> students = new ArrayList<>(studentCount);
        for (int i = 0;i<studentCount;i++){
            String name = "学生" + (i+1);
            int age = Math.abs(random.nextInt(100));
            String gender = random.nextBoolean()?"男":"女";
            Student student = new Student(name,age,gender);
            student.setId(1L+i);
            student.setPassword("abc"+i);
            student.setWeight(random.nextFloat() * 100 + 50);
            students.add(student);
        }
        return students;
    }
}