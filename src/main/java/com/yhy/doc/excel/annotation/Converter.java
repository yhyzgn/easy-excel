package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.ers.ExcelConverter;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 15:05
 * version: 1.0.0
 * desc   : 字段值类型转换器
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Converter {

    /**
     * 转换器
     *
     * @return 具体的转换器
     */
    Class<? extends ExcelConverter> value();
}
