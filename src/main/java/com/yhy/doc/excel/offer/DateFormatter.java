package com.yhy.doc.excel.offer;

import com.yhy.doc.excel.ers.ExcelFormatter;
import com.yhy.doc.excel.utils.StringUtils;

import java.util.Date;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 9:21
 * version: 1.0.0
 * desc   : 默认的日期格式化转换器
 */
public class DateFormatter implements ExcelFormatter<Date> {
    @Override
    public Date read(Object value) throws Exception {
        if (null == value) return null;
        if (value instanceof Date) {
            return (Date) value;
        }
        if (value instanceof Number) {
            return new Date(((Number) value).longValue());
        }
        if (value instanceof String) {
            String temp = (String) value;
            if (StringUtils.isNumber(temp)) {
                return new Date(Long.parseLong(temp));
            }
        }
        return null;
    }

    @Override
    public Object write(Date value) throws Exception {
        return value;
    }
}
