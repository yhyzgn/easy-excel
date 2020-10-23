package com.yhy.doc.excel.offer;

import com.yhy.doc.excel.internal.EDateParser;
import com.yhy.doc.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.DateUtil;

import java.sql.Timestamp;
import java.util.Date;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 10:35
 * version: 1.0.0
 * desc   : Timestamp 格式化
 */
public class TimestampParser implements EDateParser<Object, Timestamp> {

    @Override
    public Timestamp parse(Object value) {
        if (null == value) return null;
        if (value instanceof Timestamp) {
            return (Timestamp) value;
        }
        if (value instanceof Date) {
            return Timestamp.from(((Date) value).toInstant());
        }
        if (value instanceof Number) {
            return Timestamp.from(new Date(((Number) value).longValue()).toInstant());
        }
        if (value instanceof String) {
            String temp = (String) value;
            if (StringUtils.isNumber(temp)) {
                return Timestamp.from(new Date(Long.parseLong(temp)).toInstant());
            }
            return Timestamp.from(new Date(DateUtil.parseYYYYMMDDDate(temp).getTime()).toInstant());
        }
        return null;
    }
}
