package xyz.wqf31415.entity;

/**
 * Created by Administrator on 2019/7/12.
 *
 * @author WeiQuanfu
 */
public class Clazz {
    private String name;
    private Teacher teacher;

    public Clazz() {
    }

    public Clazz(String name, Teacher teacher) {
        this.name = name;
        this.teacher = teacher;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Teacher getTeacher() {
        return teacher;
    }

    public void setTeacher(Teacher teacher) {
        this.teacher = teacher;
    }

    @Override
    public String toString() {
        return "{" +
                "\"name\":\"" + name + '\"' +
                ", \"teacher\":" + teacher +
                '}';
    }
}
