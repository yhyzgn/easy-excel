package com.yhy.doc.excel.offer;

import com.yhy.doc.excel.ers.ExcelFormatter;
import com.yhy.doc.excel.utils.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;

import java.sql.Timestamp;
import java.util.Date;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 10:35
 * version: 1.0.0
 * desc   : Timestamp 格式化
 */
public class TimestampFormatter implements ExcelFormatter<Timestamp> {
    @Override
    public Timestamp read(Object value) throws Exception {
        if (null == value) return null;
        if (value instanceof Timestamp) {
            return (Timestamp) value;
        }
        if (value instanceof Number) {
            return Timestamp.from(new Date(((Number) value).longValue()).toInstant());
        }
        if (value instanceof String) {
            String temp = (String) value;
            if (StringUtils.isNumber(temp)) {
                return Timestamp.from(new Date(Long.parseLong(temp)).toInstant());
            }
            return Timestamp.from(new Date(HSSFDateUtil.parseYYYYMMDDDate(temp).getTime()).toInstant());
        }
        return null;
    }

    @Override
    public Object write(Timestamp value) throws Exception {
        return Date.from(value.toInstant());
    }
}
