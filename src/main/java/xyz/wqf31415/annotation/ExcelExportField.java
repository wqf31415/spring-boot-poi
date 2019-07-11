package xyz.wqf31415.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by Administrator on 2019/7/11.
 *
 * @author WeiQuanfu
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(value = ElementType.FIELD)
public @interface ExcelExportField {
    /**
     * 字段标题
     * @return
     */
    String title() default "";

    /**
     * 顺序
     * 数值小的，排在前面
     * @return
     */
    int order() default  Integer.MAX_VALUE;

    /**
     * 忽略
     * 导出到 excel 时是否忽略此字段
     * @return
     */
    boolean ignore() default false;
}
