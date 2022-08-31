package com.yhy.doc.excel.annotation;

import java.lang.annotation.*;

/**
 * 文档信息
 * <p>
 * Created on 2019-05-02 16:51
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Document {

    /**
     * 标题样式
     *
     * @return 标题样式
     */
    Style titleStyle() default @Style(
        font = @Font(
            size = 12,
            bold = true
        )
    );
}
