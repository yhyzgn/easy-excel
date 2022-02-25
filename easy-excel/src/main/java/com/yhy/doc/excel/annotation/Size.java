package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.internal.EConstant;

import java.lang.annotation.*;

/**
 * 单元格尺寸
 * <p>
 * Created on 2019-05-02 2:28
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
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
