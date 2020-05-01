package com.yhy.doc.excel.internal;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-01 10:20 下午
 * version: 1.0.0
 * desc   : 日期解析器
 */
public interface EDateParser<S, T> {

    /**
     * 解析日期格式
     *
     * @param s 原数据
     * @return 目标类型
     */
    T parse(S s) throws Exception;
}
