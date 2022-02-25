package com.yhy.doc.excel.of;

import com.yhy.doc.excel.internal.EDateParser;
import com.yhy.doc.excel.utils.StringUtils;

import java.util.Date;

/**
 * 默认的日期格式化转换器
 * <p>
 * Created on 2019-09-10 9:21
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
public class DateParser implements EDateParser<Object, Date> {

    @Override
    public Date parse(Object value) {
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
}
