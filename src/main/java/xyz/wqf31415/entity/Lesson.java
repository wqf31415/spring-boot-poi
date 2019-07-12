package xyz.wqf31415.entity;

/**
 * Created by Administrator on 2019/7/12.
 *
 * @author WeiQuanfu
 */
public class Lesson {
    private String title;

    private Float score;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Float getScore() {
        return score;
    }

    public void setScore(Float score) {
        this.score = score;
    }

    public Lesson() {
    }

    public Lesson(String title, Float score) {
        this.title = title;
        this.score = score;
    }

    @Override
    public String toString() {
        return "{" +
                "\"title\":\"" + title + '\"' +
                ", \"score\":" + score +
                '}';
    }
}
