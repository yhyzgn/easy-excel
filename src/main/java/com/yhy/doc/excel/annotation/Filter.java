package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.ers.ExcelFilter;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 15:03
 * version: 1.0.0
 * desc   : 字段过滤器
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Filter {

    /**
     * 过滤器
     *
     * @return 具体的过滤器
     */
    Class<? extends ExcelFilter> value();
}
