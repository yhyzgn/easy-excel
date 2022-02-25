package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.internal.EConstant;

import java.lang.annotation.*;

/**
 * 数据格式化模式
 * <p>
 * Created on 2019-05-01 22:40
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Pattern {

    /**
     * 格式
     *
     * @return 格式
     */
    String value() default EConstant.PATTERN_DATE_TIME;
}
