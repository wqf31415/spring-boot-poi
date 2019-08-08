package xyz.wqf31415.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by Administrator on 2019/7/15.
 *
 * @author WeiQuanfu
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelImportField {
    String title() default "";

    /**
     * 是否忽略
     * @return
     */
    boolean ignore() default false;

    /**
     * 是否必须
     * @return
     */
    boolean require() default false;

    /**
     * 是否可为空
     * @return
     */
    boolean nullable() default true;

    /**
     * 时间对象格式化
     * @return
     */
    String formatter() default "yyyy-MM-dd HH:mm:ss";
}
