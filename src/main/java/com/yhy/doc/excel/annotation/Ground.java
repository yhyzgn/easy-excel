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
     * 纹理颜色
     *
     * @return 纹理颜色
     * @see IndexedColors
     */
    IndexedColors fore();

    /**
     * 背景颜色
     *
     * @return 背景颜色
     * @see IndexedColors
     */
    IndexedColors back();

    /**
     * 纹理填充风格
     *
     * @return 纹理填充风格
     * @see FillPatternType
     */
    FillPatternType pattern();
}
