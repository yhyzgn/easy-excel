package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.internal.EFilter;

import java.lang.annotation.*;

/**
 * 字段过滤器
 * <p>
 * Created on 2019-09-09 15:03
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
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
    Class<? extends EFilter<?>> value();
}
