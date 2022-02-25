package com.yhy.doc.excel.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

/**
 * 文档信息
 * <p>
 * Created on 2019-05-02 16:51
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
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
