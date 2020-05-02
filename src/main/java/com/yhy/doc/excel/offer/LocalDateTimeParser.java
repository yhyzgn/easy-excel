package com.yhy.doc.excel.offer;

import com.yhy.doc.excel.internal.EDateParser;
import com.yhy.doc.excel.utils.ExcelUtils;
import com.yhy.doc.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.DateUtil;

import java.time.LocalDateTime;
import java.util.Date;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 10:18
 * version: 1.0.0
 * desc   : 日期时间格式化
 */
public class LocalDateTimeParser implements EDateParser<Object, LocalDateTime> {

    @Override
    public LocalDateTime parse(Object value) {
        if (null == value) return null;
        if (value instanceof Date) {
            Date date = (Date) value;
            return ExcelUtils.convertDate(date);
        }
        if (value instanceof LocalDateTime) {
            return (LocalDateTime) value;
        }
        if (value instanceof Number) {
            return ExcelUtils.convertDate(new Date(((Number) value).longValue()));
        }
        if (value instanceof String) {
            String temp = (String) value;
            if (StringUtils.isNumber(temp)) {
                return ExcelUtils.convertDate(new Date(Long.parseLong(temp)));
            }
            return ExcelUtils.convertDate(DateUtil.parseYYYYMMDDDate(temp));
        }
        return null;
    }
}
