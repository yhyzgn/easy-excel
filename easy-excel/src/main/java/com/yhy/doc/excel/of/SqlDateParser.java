package com.yhy.doc.excel.of;

import com.yhy.doc.excel.internal.EDateParser;
import com.yhy.doc.excel.utils.StringUtils;
import org.apache.poi.ss.usermodel.DateUtil;

import java.sql.Date;

/**
 * SQL 日期类格式化
 * <p>
 * Created on 2019-09-10 10:32
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
public class SqlDateParser implements EDateParser<Object, Date> {

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
            return new Date(DateUtil.parseYYYYMMDDDate(temp).getTime());
        }
        return null;
    }
}
