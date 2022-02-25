package com.yhy.doc.excel.of;

import com.yhy.doc.excel.internal.EDateParser;
import com.yhy.doc.excel.utils.ExcelUtils;
import com.yhy.doc.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.DateUtil;

import java.time.LocalDateTime;
import java.util.Date;

/**
 * 日期时间格式化
 * <p>
 * Created on 2019-09-10 10:18
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
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
