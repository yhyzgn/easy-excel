package com.yhy.doc.excel.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-02 4:51 下午
 * version: 1.0.0
 * desc   : 文档信息
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Document {

    /**
     * 标题样式
     *
     * @return 标题样式
     */
    Style titleStyle() default @Style(
            align = @Align,
            border = @Border,
            font = @Font(
                    size = 14,
                    bold = true
            ),
            ground = @Ground(
                    back = IndexedColors.GREY_25_PERCENT,
                    fore = IndexedColors.GREY_25_PERCENT
            ),
            size = @Size
    );
}
