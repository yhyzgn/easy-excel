package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.internal.EDateParser;

import java.lang.annotation.*;

/**
 * 字段值的格式化程序
 * <p>
 * Created on 2019-09-09 14:32
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Parser {

    /**
     * 格式化的实现类
     *
     * @return 格式化的实现类
     */
    Class<? extends EDateParser<?, ?>> value();
}
