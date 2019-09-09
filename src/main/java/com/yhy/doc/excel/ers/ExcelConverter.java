package com.yhy.doc.excel.ers;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 14:59
 * version: 1.0.0
 * desc   : 字段值转换器
 */
public interface ExcelConverter<E, M> {

    /**
     * 读取转换
     *
     * @param excel Excel中该字段的值
     * @return 转换到Model中该字段的值
     */
    M read(E excel);

    /**
     * @param model Model中该字段的值
     * @return 转换到Excel中该字段的值
     */
    E write(M model);
}
