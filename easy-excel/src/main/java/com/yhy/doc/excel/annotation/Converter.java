package com.yhy.doc.excel.annotation;

import com.yhy.doc.excel.internal.EConverter;

import java.lang.annotation.*;

/**
 * 字段值类型转换器
 * <p>
 * Created on 2019-09-09 15:05
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
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
    Class<? extends EConverter<?, ?>> value();
}
