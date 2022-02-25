package com.yhy.doc.excel.annotation;

import java.lang.annotation.*;

/**
 * 单元格样式风格
 * <p>
 * Created on 2019-05-02 1:30
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Style {
    /**
     * 单元格对齐方式
     *
     * @return 单元格对齐方式
     */
    Align align() default @Align;

    /**
     * 边框样式
     *
     * @return 边框样式
     */
    Border border() default @Border;

    /**
     * 字体样式
     *
     * @return 字体样式
     */
    Font font() default @Font;

    /**
     * 背景和纹理
     *
     * @return 背景和纹理
     */
    Ground ground() default @Ground;

    /**
     * 边框大小
     *
     * @return 边框大小
     */
    Size size() default @Size;
}
