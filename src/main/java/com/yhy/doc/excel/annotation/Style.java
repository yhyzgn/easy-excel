package com.yhy.doc.excel.annotation;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-02 1:30 上午
 * version: 1.0.0
 * desc   : 单元格样式风格
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
    Align align();

    /**
     * 边框样式
     *
     * @return 边框样式
     */
    Border border();

    /**
     * 字体样式
     *
     * @return 字体样式
     */
    Font font();

    /**
     * 背景和纹理
     *
     * @return 背景和纹理
     */
    Ground ground();

    /**
     * 边框大小
     *
     * @return 边框大小
     */
    Size size();
}
