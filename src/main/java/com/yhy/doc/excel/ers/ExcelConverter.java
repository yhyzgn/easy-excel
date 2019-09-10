package com.yhy.doc.excel.ers;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 14:59
 * version: 1.0.0
 * desc   : 字段值转换器
 */
public interface ExcelConverter<T> {

    /**
     * 读取转换
     *
     * @param value Excel中该字段的值
     * @return 转换到Model中该字段的值
     */
    T read(Object value);

    /**
     * @param value Model中该字段的值
     * @return 转换到Excel中该字段的值
     */
    Object write(T value);
}
