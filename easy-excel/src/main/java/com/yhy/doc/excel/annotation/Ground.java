package com.yhy.doc.excel.annotation;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

/**
 * 背景和纹理
 * <p>
 * Created on 2019-05-02 2:22
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
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
