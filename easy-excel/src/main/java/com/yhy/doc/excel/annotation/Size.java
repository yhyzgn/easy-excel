package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.internal.EConstant;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-02 2:28 上午
 * version: 1.0.0
 * desc   : 单元格尺寸
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Size {

    /**
     * 是否启用
     *
     * @return 是否启用
     */
    boolean enabled() default true;

    /**
     * 宽度
     *
     * @return 宽度
     */
    int width() default EConstant.COLUMN_WIDTH;

    /**
     * 高度
     *
     * @return 高度
     */
    short height() default EConstant.ROW_HEIGHT;
}
