package xyz.wqf31415.entity;

/**
 * Created by Administrator on 2019/7/12.
 *
 * @author WeiQuanfu
 */
public class Teacher extends Human{
    private String name;

    private String phone;

    private String subject;

    public Teacher() {
    }

    public Teacher(String name, String phone, String subject) {
        this.name = name;
        this.phone = phone;
        this.subject = subject;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getSubject() {
        return subject;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    @Override
    public String toString() {
        return "{" +
                "\"name\":\"" + name + '\"' +
                ", \"phone\":\"" + phone + '\"' +
                ", \"subject\":\"" + subject + '\"' +
                '}';
    }
}
