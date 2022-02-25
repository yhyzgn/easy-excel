package com.yhy.doc.excel.annotation;

import java.lang.annotation.*;

/**
 * 导出时字段名称排序
 * <p>
 * Created on 2019-04-25 20:05
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
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
