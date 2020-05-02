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
     * 水平对齐方式
     *
     * @return 水平对齐方式
     * @see HorizontalAlignment
     */
    HorizontalAlignment horizontal();

    /**
     * 垂直对齐方式
     *
     * @return 垂直对齐方式
     * @see VerticalAlignment
     */
    VerticalAlignment vertical();
}
