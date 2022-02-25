package com.yhy.doc.excel.annotation;

import java.lang.annotation.*;

/**
 * 导出时忽略当前字段
 * <p>
 * Created on 2019-04-24 21:42
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Ignored {
}
