package com.yhy.doc.excel;

import com.yhy.doc.excel.entity.Sex;
import com.yhy.doc.excel.ers.ExcelConverter;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 16:14
 * version: 1.0.0
 * desc   :
 */
public class SexConverter implements ExcelConverter<Sex> {
    @Override
    public Sex read(Object value) {
        if (value instanceof String) {
            return Sex.parse((String) value);
        }
        return null;
    }

    @Override
    public Object write(Sex value) {
        return value.getValue();
    }
}
