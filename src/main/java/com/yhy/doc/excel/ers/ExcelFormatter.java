package com.yhy.doc.excel.ers;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 14:36
 * version: 1.0.0
 * desc   : 格式化转换器接口
 */
public interface ExcelFormatter<T> {

    /**
     * 读取时格式化
     *
     * @param value 原始值
     * @return 格式化后的值
     */
    T read(String value) throws Exception;

    /**
     * 写数据时格式化
     *
     * @param value 原始值
     * @return 格式化后的值
     */
    String write(T value) throws Exception;
}
