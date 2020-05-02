package com.yhy.doc.excel.utils;

import com.yhy.doc.excel.annotation.Converter;
import com.yhy.doc.excel.annotation.Filter;
import com.yhy.doc.excel.annotation.Parser;
import com.yhy.doc.excel.compat.GetterSetter;
import com.yhy.doc.excel.extra.ExcelColumn;
import com.yhy.doc.excel.extra.ReaderConfig;
import com.yhy.doc.excel.extra.Rect;
import com.yhy.doc.excel.internal.EConverter;
import com.yhy.doc.excel.internal.EDateParser;
import com.yhy.doc.excel.internal.EFilter;
import com.yhy.doc.excel.io.ExcelReader;
import com.yhy.doc.excel.io.ExcelWriter;
import com.yhy.doc.excel.offer.DateParser;
import com.yhy.doc.excel.offer.LocalDateTimeParser;
import com.yhy.doc.excel.offer.SqlDateParser;
import com.yhy.doc.excel.offer.TimestampParser;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.ServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.*;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatterBuilder;
import java.util.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:14
 * version: 1.0.0
 * desc   : 工具类
 */
@Slf4j
public class ExcelUtils {

    private ExcelUtils() {
        throw new UnsupportedOperationException("Utils class can not be instantiate.");
    }

    public static <T> List<T> read(File file, Class<T> clazz) {
        return read(file, ReaderConfig.deft(), clazz);
    }

    public static <T> List<T> read(File file, ReaderConfig config, Class<T> clazz) {
        try {
            return read(new FileInputStream(file), config, clazz);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    public static <T> List<T> read(InputStream is, Class<T> clazz) {
        return read(is, ReaderConfig.deft(), clazz);
    }

    public static <T> List<T> read(InputStream is, ReaderConfig config, Class<T> clazz) {
        ExcelReader<T> reader = null;
        try {
            return new ExcelReader<T>(is, config).read(clazz);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    public static <T> List<T> read(ServletRequest request, Class<T> clazz) {
        return read(request, ReaderConfig.deft(), clazz);
    }

    public static <T> List<T> read(ServletRequest request, ReaderConfig config, Class<T> clazz) {
        try {
            return read(request.getInputStream(), config, clazz);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    public static <T> void write(File file, List<T> src) {
        write(file, src, defaultSheet());
    }

    public static <T> void write(File file, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(file).write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static <T> void writeX(File file, List<T> src) {
        writeX(file, src, defaultSheet());
    }

    public static <T> void writeX(File file, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(file).x().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static <T> void writeBig(File file, List<T> src) {
        writeBig(file, src, defaultSheet());
    }

    public static <T> void writeBig(File file, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(file).x().big().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static <T> void write(OutputStream os, List<T> src) {
        write(os, src, defaultSheet());
    }

    public static <T> void write(OutputStream os, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(os).write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static <T> void writeX(OutputStream os, List<T> src) {
        writeX(os, src, defaultSheet());
    }

    public static <T> void writeX(OutputStream os, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(os).x().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static <T> void writeBig(OutputStream os, List<T> src) {
        writeBig(os, src, defaultSheet());
    }

    public static <T> void writeBig(OutputStream os, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(os).x().big().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static <T> void write(HttpServletResponse response, List<T> src) {
        write(response, src, defaultSheet());
    }

    public static <T> void write(HttpServletResponse response, List<T> src, String sheetName) {
        write(response, defaultFilename(), src, sheetName);
    }

    public static <T> void writeX(HttpServletResponse response, List<T> src) {
        writeX(response, src, defaultSheet());
    }

    public static <T> void writeX(HttpServletResponse response, List<T> src, String sheetName) {
        writeX(response, defaultFilename(), src, sheetName);
    }

    public static <T> void writeBig(HttpServletResponse response, List<T> src) {
        writeBig(response, src, defaultSheet());
    }

    public static <T> void writeBig(HttpServletResponse response, List<T> src, String sheetName) {
        writeBig(response, defaultFilename(), src, sheetName);
    }

    public static <T> void write(HttpServletResponse response, String filename, List<T> src) {
        write(response, filename, src, defaultSheet());
    }

    public static <T> void write(HttpServletResponse response, String filename, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(response, filename).write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static <T> void writeX(HttpServletResponse response, String filename, List<T> src) {
        writeX(response, filename, src, defaultSheet());
    }

    public static <T> void writeX(HttpServletResponse response, String filename, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(response, filename).x().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static <T> void writeBig(HttpServletResponse response, String filename, List<T> src) {
        writeBig(response, filename, src, defaultSheet());
    }

    public static <T> void writeBig(HttpServletResponse response, String filename, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(response, filename).x().big().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String defaultFilename() {
        return new SimpleDateFormat("yyyy-MM-dd_HH_mm_ss").format(Calendar.getInstance(Locale.getDefault()).getTime()) + ".xlsx";
    }

    public static String defaultSheet() {
        return "Sheet1";
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

    public static String formatDate(Date date, String pattern) {
        return date.toInstant().atZone(ZoneId.systemDefault()).format(new DateTimeFormatterBuilder().appendPattern(pattern).toFormatter(Locale.getDefault()));
    }

    public static String formatDate(LocalDateTime time, String pattern) {
        return time.format(new DateTimeFormatterBuilder().appendPattern(pattern).toFormatter(Locale.getDefault()));
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

    public static void invokeSetter(Object obj, Field field, Object value) throws Exception {
        GetterSetter.invokeSetter(obj, field, value);
    }

    public static Object invokeGetter(Object obj, Field field) throws Exception {
        return GetterSetter.invokeGetter(obj, field);
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

    public static DateParser offeredDateParser() {
        return new DateParser();
    }

    public static LocalDateTimeParser offeredLocalDateTimeParser() {
        return new LocalDateTimeParser();
    }

    public static SqlDateParser offeredSqlDateParser() {
        return new SqlDateParser();
    }

    public static TimestampParser offeredTimestampParser() {
        return new TimestampParser();
    }

    public static void checkColumn(ExcelColumn title, Field field) {
        // 过滤器
        Filter filter = field.getAnnotation(Filter.class);
        if (null != filter && filter.value() != EFilter.class) {
            title.setFilter(instantiate(filter.value()));
        }
        // 类型转换器
        Converter converter = field.getAnnotation(Converter.class);
        if (null != converter && converter.value() != EConverter.class) {
            title.setConverter(instantiate(converter.value()));
        }
        // 日期解析器
        Parser parser = field.getAnnotation(Parser.class);
        if (null != parser && parser.value() != EDateParser.class) {
            title.setParser(instantiate(parser.value()));
        }
    }
}
