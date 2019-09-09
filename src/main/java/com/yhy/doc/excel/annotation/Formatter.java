package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.ers.ExcelFormatter;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 14:32
 * version: 1.0.0
 * desc   : 字段值的格式化程序
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Formatter {

    /**
     * 格式化的实现类
     *
     * @return 格式化的实现类
     */
    Class<? extends ExcelFormatter> value();
}
