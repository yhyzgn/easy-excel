package com.yhy.doc.excel.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-02 1:36 上午
 * version: 1.0.0
 * desc   : 单元格对齐方式
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Align {

    /**
     * 是否启用
     *
     * @return 是否启用
     */
    boolean enabled() default true;

    /**
     * 水平对齐方式
     *
     * @return 水平对齐方式
     * @see HorizontalAlignment
     */
    HorizontalAlignment horizontal() default HorizontalAlignment.CENTER;

    /**
     * 垂直对齐方式
     *
     * @return 垂直对齐方式
     * @see VerticalAlignment
     */
    VerticalAlignment vertical() default VerticalAlignment.CENTER;

    /**
     * 是否自动换行
     *
     * @return 是否自动换行
     */
    boolean wrap() default true;

    /**
     * 缩进
     *
     * @return 缩进
     */
    short indention() default 0;

    /**
     * 文本旋转
     *
     * @return 文本旋转，取值是从-90到90，而不是0-180度
     */
    short rotation() default 0;
}
