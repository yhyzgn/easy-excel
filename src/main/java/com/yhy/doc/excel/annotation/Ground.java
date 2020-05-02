package com.yhy.doc.excel.annotation;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-02 2:22 上午
 * version: 1.0.0
 * desc   : 背景和纹理
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Ground {

    /**
     * 是否启用
     *
     * @return 是否启用
     */
    boolean enabled() default true;

    /**
     * 纹理颜色
     *
     * @return 纹理颜色
     * @see IndexedColors
     */
    IndexedColors fore() default IndexedColors.WHITE;

    /**
     * 背景颜色
     *
     * @return 背景颜色
     * @see IndexedColors
     */
    IndexedColors back() default IndexedColors.WHITE;

    /**
     * 设置图案样式
     *
     * @return 设置图案样式
     * @see FillPatternType
     */
    FillPatternType pattern() default FillPatternType.SOLID_FOREGROUND;
}
