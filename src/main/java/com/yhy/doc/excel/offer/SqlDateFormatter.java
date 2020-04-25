package com.yhy.doc.excel.offer;

import com.yhy.doc.excel.internal.ExcelConstant;
import com.yhy.doc.excel.internal.ExcelFormatter;
import com.yhy.doc.excel.utils.ExcelUtils;
import com.yhy.doc.excel.utils.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;

import java.sql.Date;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 10:32
 * version: 1.0.0
 * desc   : SQL日期类格式化
 */
public class SqlDateFormatter implements ExcelFormatter<Object, Date> {
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
            return new Date(HSSFDateUtil.parseYYYYMMDDDate(temp).getTime());
        }
        return null;
    }

    @Override
    public Object write(Date value) throws Exception {
        return ExcelUtils.formatDate(value, ExcelConstant.PATTERN);
    }
}
