package com.yhy.doc.excel.utils;

import com.yhy.doc.excel.internal.ReaderConfig;
import com.yhy.doc.excel.internal.Rect;
import com.yhy.doc.excel.io.ExcelReader;
import com.yhy.doc.excel.offer.DateFormatter;
import com.yhy.doc.excel.offer.LocalDateTimeFormatter;
import com.yhy.doc.excel.offer.SqlDateFormatter;
import com.yhy.doc.excel.offer.TimestampFormatter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.InputStream;
import java.lang.reflect.*;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Collections;
import java.util.Date;
import java.util.List;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:14
 * version: 1.0.0
 * desc   :
 */
@Slf4j
public class ExcelUtils {

    private ExcelUtils() {
        throw new UnsupportedOperationException("Utils class can not be instantiate.");
    }

    public static <T> List<T> read(File file, Class<T> clazz) {
        return read(file, null, clazz);
    }

    public static <T> List<T> read(File file, ReaderConfig config, Class<T> clazz) {
        ExcelReader<T> reader = ExcelReader.create(file, config);
        if (null != reader) {
            return reader.read(clazz);
        }
        return Collections.emptyList();
    }

    public static <T> List<T> read(InputStream is, Class<T> clazz) {
        return read(is, null, clazz);
    }

    public static <T> List<T> read(InputStream is, ReaderConfig config, Class<T> clazz) {
        ExcelReader<T> reader = ExcelReader.create(is, config);
        return reader.read(clazz);
    }

    public static Rect merged(Sheet sheet, int row, int column, int rowStartIndex, int columnStartIndex) {
        int mergedCount = sheet.getNumMergedRegions();
        CellRangeAddress range;
        int firstColumn, lastColumn, firstRow, lastRow;
        for (int i = 0; i < mergedCount; i++) {
            range = sheet.getMergedRegion(i);
            firstRow = range.getFirstRow() - rowStartIndex;
            lastRow = range.getLastRow() - rowStartIndex;
            firstColumn = range.getFirstColumn() - columnStartIndex;
            lastColumn = range.getLastColumn() - columnStartIndex;
            if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
                return new Rect(true, firstRow, lastRow, firstColumn, lastColumn);
            }
        }
        return new Rect(false, row, row, column, column);
    }

    public static LocalDateTime convertDate(Date date) {
        return LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault());
    }

    public static Date convertDate(LocalDateTime time) {
        return Date.from(time.atZone(ZoneId.systemDefault()).toInstant());
    }

    public static <T> T instantiate(Class<T> clazz) {
        try {
            return instantiate(clazz.getConstructor());
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static <T> T instantiate(Constructor<T> constructor) {
        try {
            return constructor.newInstance();
        } catch (InstantiationException | IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static Method setter(Field field, Class<?> clazz) throws IntrospectionException {
        PropertyDescriptor descriptor = new PropertyDescriptor(field.getName(), clazz);
        return descriptor.getWriteMethod();
    }

    public static Method getter(Field field, Class<?> clazz) throws IntrospectionException {
        PropertyDescriptor descriptor = new PropertyDescriptor(field.getName(), clazz);
        return descriptor.getReadMethod();
    }

    public static Type getParamType(Class<?> clazz, Class<?> interfaceClazz, int index) {
        Type[] types = clazz.getGenericInterfaces();
        ParameterizedType ptp;
        Type raw;
        for (Type tp : types) {
            if (tp instanceof ParameterizedType) {
                ptp = (ParameterizedType) tp;
                raw = ptp.getRawType();
                if (raw.getTypeName().equalsIgnoreCase(interfaceClazz.getName()) && ptp.getActualTypeArguments().length > index) {
                    // 匹配到对应的接口
                    return ptp.getActualTypeArguments()[index];
                }
            }
        }
        return null;
    }

    public static DateFormatter offeredDateFormatter() {
        return new DateFormatter();
    }

    public static LocalDateTimeFormatter offeredLocalDateTimeFormatter() {
        return new LocalDateTimeFormatter();
    }

    public static SqlDateFormatter offeredSqlDateFormatter() {
        return new SqlDateFormatter();
    }

    public static TimestampFormatter offeredTimestampFormatter() {
        return new TimestampFormatter();
    }
}
