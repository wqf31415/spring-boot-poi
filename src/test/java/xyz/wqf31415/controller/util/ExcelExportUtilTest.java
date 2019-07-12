package xyz.wqf31415.controller.util;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;
import xyz.wqf31415.annotation.ExcelExportField;
import xyz.wqf31415.entity.Clazz;
import xyz.wqf31415.entity.Lesson;
import xyz.wqf31415.entity.Student;
import xyz.wqf31415.entity.Teacher;

import java.io.File;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.ZoneId;
import java.time.ZonedDateTime;
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
    public void exportData() throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        Random random = new Random();
        int studentListCount = random.nextInt(4)+1;
        for (int i = 0;i<studentListCount;i++){
            ExcelExportUtil.generateSheetAndWriteData(workbook, generateStudentList(),Student.class,"student_"+(i+1),"学生名单("+(i+1)+')');
        }
        ExcelExportUtil.generateSheetAndWriteData(workbook,generateTeacherList(),Teacher.class,"teacher","教师名单");

        // 写入文件
        File file = new File(exportFilePath);
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
    }

    private List<Student> generateStudentList(){
        Random random = new Random();
        int studentCount = random.nextInt(51)+20;
        List<Student> students = new ArrayList<>(studentCount);
        Teacher teacher = new Teacher(NameUtil.lastNameArray[random.nextInt(NameUtil.lastNameArray.length)] +NameUtil.firstNameArray[random.nextInt(NameUtil.firstNameArray.length)],"13512341234","课程"+(random.nextInt(3)+1));
        Clazz clazz = new Clazz("班级"+(random.nextInt(3)+1),teacher);
        for (int i = 0;i<studentCount;i++){
            String name = NameUtil.lastNameArray[random.nextInt(NameUtil.lastNameArray.length)] + NameUtil.firstNameArray[random.nextInt(NameUtil.firstNameArray.length)];
            int year = random.nextInt(20)+1980;
            int age = 2019 - year;
            String gender = random.nextBoolean()?"男":"女";
            Student student = new Student(name,age,gender);
            student.setId(1L+i);
            student.setPassword("abc"+i);
            student.setWeight(random.nextFloat() * 60 + 50);
            student.setBirth(ZonedDateTime.of(year,random.nextInt(11)+1,random.nextInt(27)+1,random.nextInt(24),random.nextInt(60),random.nextInt(60),0, ZoneId.systemDefault()));
            student.setClazz(clazz);
            student.setLessons(generateLessonList());
            students.add(student);
        }
        return students;
    }

    private List<Lesson> generateLessonList(){
        List<Lesson> lessons = new ArrayList<>();
        Random random = new Random();
        int count = random.nextInt(5)+1;
        for (int i=0;i<count;i++){
            lessons.add(new Lesson("课程"+(i+1),Math.round(random.nextFloat()*1000)/10f));
        }
        return lessons;
    }

    private List<Teacher> generateTeacherList(){
        Random random = new Random();
        int count = random.nextInt(40)+10;
        List<Teacher> teachers = new ArrayList<>(count);
        for (int i = 0;i<count;i++){
            teachers.add(new Teacher(generateRandomName(random),"13","科目"));
        }
        return teachers;
    }

    private String generateRandomName(Random random){
        return NameUtil.lastNameArray[random.nextInt(NameUtil.lastNameArray.length)] +NameUtil.firstNameArray[random.nextInt(NameUtil.firstNameArray.length)];
    }
}