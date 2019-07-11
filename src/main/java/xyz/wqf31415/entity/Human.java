package xyz.wqf31415.entity;

import xyz.wqf31415.annotation.ExcelExportField;

/**
 * Created by Administrator on 2019/7/11.
 *
 * @author WeiQuanfu
 */
public class Human extends Animal{
    @ExcelExportField(order = 0,title = "ID")
    private Long id;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }
}
