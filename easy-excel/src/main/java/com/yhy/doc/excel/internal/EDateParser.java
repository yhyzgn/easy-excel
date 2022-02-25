package com.yhy.doc.excel.internal;

/**
 * 日期解析器
 * <p>
 * Created on 2019-09-09 22:20
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
public interface EDateParser<S, T> {

    /**
     * 解析日期格式
     *
     * @param s 原数据
     * @return 目标类型
     * @throws Exception 可能会出现的异常
     */
    T parse(S s) throws Exception;
}
