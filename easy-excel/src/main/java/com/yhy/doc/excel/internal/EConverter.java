package com.yhy.doc.excel.internal;

/**
 * 字段值转换器
 * <p>
 * Created on 2019-09-09 14:59
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
public interface EConverter<E, M> {

    /**
     * 读取转换
     *
     * @param value Excel中该字段的值
     * @return 转换到Model中该字段的值
     */
    M read(E value);

    /**
     * @param value Model中该字段的值
     * @return 转换到Excel中该字段的值
     */
    E write(M value);
}
