package com.yhy.doc.excel.annotation;

import java.lang.annotation.*;

/**
 * 字段注解
 * <p>
 * Created on 2019-09-09 14:04
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Column {

    /**
     * 字段名称
     *
     * @return 字段名称
     */
    String value() default "";

    /**
     * 是否可空
     *
     * @return 是否可空
     */
    boolean nullable() default true;

    /**
     * 模糊匹配字段名称模板，如：%名称%
     *
     * @return 模糊匹配模板
     */
    String like() default "";

    /**
     * 是否智能匹配，采用字符串相似度匹配
     *
     * @return 是否智能匹配
     */
    boolean intelligent() default false;

    /**
     * 智能匹配容差，容错率
     * <p>
     * 只有相似度 ≥ (1 - tolerance) 才能匹配成功
     *
     * @return 智能匹配容错率
     */
    double tolerance() default 0.4;

    /**
     * 是否自动处理换行符
     *
     * @return 是否自动处理换行符
     */
    boolean wrap() default false;

    /**
     * 是否大小写不敏感
     *
     * @return 是否大小写不敏感
     */
    boolean insensitive() default true;

    /**
     * 导出时的字段名，优先获取，不指定则获取value()
     *
     * @return 导出时的字段名
     */
    String export() default "";

    /**
     * 导出时的计算公式
     * <p>
     * 如：A1*B1，SUM(A1,C1)，LOOKUP(A5,$A$1:$A$4,$C$1:$C$4)等
     *
     * @return 导出时的计算公式
     */
    String formula() default "";
}
