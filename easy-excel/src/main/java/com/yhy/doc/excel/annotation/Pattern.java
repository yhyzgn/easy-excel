package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.internal.EConstant;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-01 10:40 下午
 * version: 1.0.0
 * desc   : 数据格式化模式
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
