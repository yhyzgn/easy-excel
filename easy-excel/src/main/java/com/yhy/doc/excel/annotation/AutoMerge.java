package com.yhy.doc.excel.annotation;

import java.lang.annotation.*;

/**
 * 自动合并单元格
 * <p>
 * Created on 2019-03-16 17:24
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface AutoMerge {
}
