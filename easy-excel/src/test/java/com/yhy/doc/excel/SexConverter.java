package com.yhy.doc.excel;

import com.yhy.doc.excel.entity.Sex;
import com.yhy.doc.excel.internal.EConverter;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 16:14
 * version: 1.0.0
 * desc   :
 */
public class SexConverter implements EConverter<String, Sex> {
    @Override
    public Sex read(String value) {
        return Sex.parse(value);
    }

    @Override
    public String write(Sex value) {
        return value.getValue();
    }
}
