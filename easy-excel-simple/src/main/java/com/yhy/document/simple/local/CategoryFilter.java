package com.yhy.document.simple.local;

import com.yhy.doc.excel.internal.EFilter;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-02 3:15 上午
 * version: 1.0.0
 * desc   :
 */
public class CategoryFilter implements EFilter<String> {

    @Override
    public String read(String value) {
        return value.endsWith("0") ? "扯淡-读" : value;
    }

    @Override
    public String write(String value) {
        return value.endsWith("2") ? "扯淡-写" : value;
    }
}
