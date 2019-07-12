package xyz.wqf31415.entity;

import xyz.wqf31415.annotation.ExcelExportField;

import java.time.ZonedDateTime;
import java.util.List;


/**
 * Created by Administrator on 2018/2/27.
 *
 * @author WeiQuanfu
 */
public class Student extends Human{
    @ExcelExportField(title = "姓名",order = 1)
    private String name;
    @ExcelExportField(title = "年龄")
    private Integer age;
    @ExcelExportField(order = 2,title = "性别")
    private String gender;
    @ExcelExportField(formatter = "yyyy-MM-dd HH:mm:ss")
    private ZonedDateTime birth;

    @ExcelExportField(ignore = true)
    private String password;

    private Clazz clazz;

    private List<Lesson> lessons;

    public Student() {
    }

    public Student(String name, Integer age, String gender) {
        this.name = name;
        this.age = age;
        this.gender = gender;
    }

    public ZonedDateTime getBirth() {
        return birth;
    }

    public void setBirth(ZonedDateTime birth) {
        this.birth = birth;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public Clazz getClazz() {
        return clazz;
    }

    public void setClazz(Clazz clazz) {
        this.clazz = clazz;
    }

    public List<Lesson> getLessons() {
        return lessons;
    }

    public void setLessons(List<Lesson> lessons) {
        this.lessons = lessons;
    }

    @Override
    public String toString() {
        return "Student{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", gender='" + gender + '\'' +
                '}';
    }
}
