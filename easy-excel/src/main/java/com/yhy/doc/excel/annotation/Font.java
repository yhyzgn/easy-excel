package com.yhy.doc.excel.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

/**
 * 字体样式
 * <p>
 * Created on 2019-05-02 2:10
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface Font {

    /**
     * 是否启用
     *
     * @return 是否启用
     */
    boolean enabled() default true;

    /**
     * 字体名称
     *
     * @return 字体名称
     */
    String name() default "微软雅黑";

    /**
     * 字号大小
     *
     * @return 字号大小
     */
    short size() default 10;

    /**
     * 是否粗体
     *
     * @return 是否粗体
     */
    boolean bold() default false;

    /**
     * 是否斜体
     *
     * @return 是否斜体
     */
    boolean italic() default false;

    /**
     * 是否有删除线
     *
     * @return 是否有删除线
     */
    boolean delete() default false;

    /**
     * 字体样色
     *
     * @return 字体样色
     */
    IndexedColors color() default IndexedColors.BLACK;

    /**
     * 下划线风格
     *
     * @return 下划线风格
     * @see org.apache.poi.ss.usermodel.Font#U_NONE
     * @see org.apache.poi.ss.usermodel.Font#U_SINGLE
     * @see org.apache.poi.ss.usermodel.Font#U_DOUBLE
     * @see org.apache.poi.ss.usermodel.Font#U_SINGLE_ACCOUNTING
     * @see org.apache.poi.ss.usermodel.Font#U_DOUBLE_ACCOUNTING
     */
    byte underline() default org.apache.poi.ss.usermodel.Font.U_NONE;

    /**
     * 设置上下标
     *
     * @return 上下标
     * @see org.apache.poi.ss.usermodel.Font#SS_NONE
     * @see org.apache.poi.ss.usermodel.Font#SS_SUPER
     * @see org.apache.poi.ss.usermodel.Font#SS_SUB
     */
    short typeOffset() default org.apache.poi.ss.usermodel.Font.SS_NONE;
}
