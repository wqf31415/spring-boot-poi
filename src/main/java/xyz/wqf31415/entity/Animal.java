package xyz.wqf31415.entity;

import xyz.wqf31415.annotation.ExcelExportField;

/**
 * Created by Administrator on 2019/7/11.
 *
 * @author WeiQuanfu
 */
public class Animal {
    /**
     * 体重
     */
    @ExcelExportField(title = "体重(kg)")
    private Float weight;

    public Float getWeight() {
        return weight;
    }

    public void setWeight(Float weight) {
        this.weight = weight;
    }
}
