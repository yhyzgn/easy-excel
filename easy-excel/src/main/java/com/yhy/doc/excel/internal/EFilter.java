package com.yhy.doc.excel.internal;

/**
 * 过滤器接口
 * <p>
 * Created on 2019-09-09 14:54
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
public interface EFilter<T> {

    /**
     * 读取过滤
     *
     * @param value 读取到的值
     * @return 过滤后的值
     */
    T read(T value);

    /**
     * 写入过滤
     *
     * @param value 写入前的值
     * @return 过滤后的值
     */
    T write(T value);
}
