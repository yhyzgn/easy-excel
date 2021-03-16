package com.yhy.doc.excel.annotation;

import java.lang.annotation.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-04-25 8:05 下午
 * version: 1.0.0
 * desc   : 导出时字段名称排序
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Sorted {

    /**
     * 列名排序序号，小优先
     *
     * @return 序号
     */
    int value();
}
