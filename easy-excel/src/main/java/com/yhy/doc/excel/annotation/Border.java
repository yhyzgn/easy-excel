package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.internal.EBorderSide;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

/**
 * 边框样式
 * <p>
 * Created on 2019-05-02 1:53
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Border {

    /**
     * 是否启用
     *
     * @return 是否启用
     */
    boolean enabled() default true;

    /**
     * 边框颜色
     *
     * @return 边框颜色
     * @see IndexedColors
     */
    IndexedColors color() default IndexedColors.BLACK;

    /**
     * 边框风格
     *
     * @return 边框风格
     * @see IndexedColors
     */
    BorderStyle style() default BorderStyle.THIN;

    /**
     * 边框方向
     *
     * @return 边框方向
     * @see EBorderSide
     */
    EBorderSide[] sides() default {EBorderSide.ALL};
}
