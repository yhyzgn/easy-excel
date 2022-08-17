package com.yhy.doc.excel.utils;

import com.yhy.doc.excel.annotation.Converter;
import com.yhy.doc.excel.annotation.Filter;
import com.yhy.doc.excel.annotation.Parser;
import com.yhy.doc.excel.extra.ExcelColumn;
import com.yhy.doc.excel.extra.Rect;
import com.yhy.doc.excel.of.DateParser;
import com.yhy.doc.excel.of.LocalDateTimeParser;
import com.yhy.doc.excel.of.SqlDateParser;
import com.yhy.doc.excel.of.TimestampParser;
import com.yhy.jakit.util.descriptor.GetterSetter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.*;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatterBuilder;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * Excel 工具类
 * <p>
 * Created on 2019-09-09 12:14
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Slf4j
public class ExcelUtils {

    /**
     * 禁止实例化处理
     */
    private ExcelUtils() {
        throw new UnsupportedOperationException("Utils class can not be instantiate.");
    }

    /**
     * 对列表过滤
     *
     * @param list      列表
     * @param predicate 过滤条件
     * @param <T>       数据类型
     * @return 过滤结果
     */
    public static <T> List<T> filter(List<T> list, Predicate<T> predicate) {
        if (null == list || null == predicate) {
            return list;
        }
        return list.stream().filter(predicate).collect(Collectors.toList());
    }

    /**
     * 默认的文件名
     *
     * @return 默认文件名
     */
    public static String defaultFilename() {
        return new SimpleDateFormat("yyyy-MM-dd_HH_mm_ss").format(Calendar.getInstance(Locale.getDefault()).getTime()) + ".xlsx";
    }

    /**
     * 默认工作簿名称
     *
     * @return 默认工作簿名称
     */
    public static String defaultSheet() {
        return "Sheet1";
    }

    /**
     * 获取合并单元格
     *
     * @param sheet  工作簿
     * @param row    当前行
     * @param column 行前列
     * @return 单元格范围
     */
    public static Rect merged(Sheet sheet, int row, int column) {
        int mergedCount = sheet.getNumMergedRegions();
        CellRangeAddress range;
        int firstColumn, lastColumn, firstRow, lastRow;
        for (int i = 0; i < mergedCount; i++) {
            range = sheet.getMergedRegion(i);
            firstRow = range.getFirstRow();
            lastRow = range.getLastRow();
            firstColumn = range.getFirstColumn();
            lastColumn = range.getLastColumn();
            if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
                return Rect.builder()
                    .merged(true)
                    .rowStart(firstRow)
                    .rowEnd(lastRow)
                    .columnStart(firstColumn)
                    .columnEnd(lastColumn)
                    .build();
            }
        }
        return Rect.builder()
            .merged(false)
            .rowStart(row)
            .rowEnd(row)
            .columnStart(column)
            .columnEnd(column)
            .build();
    }

    /**
     * Date 转换为 LocalDateTime
     *
     * @param date Date
     * @return LocalDateTime
     */
    public static LocalDateTime convertDate(Date date) {
        return LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault());
    }

    /**
     * LocalDateTime 转换为 Date
     *
     * @param time LocalDateTime
     * @return Date
     */
    public static Date convertDate(LocalDateTime time) {
        return Date.from(time.atZone(ZoneId.systemDefault()).toInstant());
    }

    /**
     * Date 格式化
     *
     * @param date    Date
     * @param pattern 格式
     * @return 结果
     */
    public static String formatDate(Date date, String pattern) {
        return date.toInstant().atZone(ZoneId.systemDefault()).format(new DateTimeFormatterBuilder().appendPattern(pattern).toFormatter(Locale.getDefault()));
    }

    /**
     * LocalDateTime 格式化
     *
     * @param time    LocalDateTime
     * @param pattern 格式
     * @return 结果
     */
    public static String formatDate(LocalDateTime time, String pattern) {
        return time.format(new DateTimeFormatterBuilder().appendPattern(pattern).toFormatter(Locale.getDefault()));
    }

    /**
     * 反射通过构造行数创建对象
     *
     * @param clazz 映射的类
     * @param <T>   映射的类
     * @return 实例
     */
    public static <T> T instantiate(Class<T> clazz) {
        try {
            return instantiate(clazz.getConstructor());
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 反射通过构造行数创建对象
     *
     * @param constructor 构造函数对象
     * @param <T>         映射的类
     * @return 实例
     */
    public static <T> T instantiate(Constructor<T> constructor) {
        try {
            return constructor.newInstance();
        } catch (InstantiationException | IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 执行字段对应的 setter 方法
     *
     * @param obj   当前对象
     * @param field 字段
     * @param value 将要被set的字段值
     * @throws Exception 可能出现的异常
     */
    public static void invokeSetter(Object obj, Field field, Object value) throws Exception {
        GetterSetter.invokeSetter(obj, field, value);
    }

    /**
     * 执行字段对应的 getter 方法
     *
     * @param obj   当前对象
     * @param field 字段
     * @return 字段get到的值
     * @throws Exception 可能出现的异常
     */
    public static Object invokeGetter(Object obj, Field field) throws Exception {
        return GetterSetter.invokeGetter(obj, field);
    }

    /**
     * 获取泛型参数
     *
     * @param clazz          当前类对象
     * @param interfaceClazz 接口对象
     * @param index          泛型索引
     * @return 泛型类型
     */
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

    /**
     * 内置的 Date 解析器
     *
     * @return 内置的 Date 解析器
     */
    public static DateParser offeredDateParser() {
        return new DateParser();
    }

    /**
     * 内置的 LocalDateTime 解析器
     *
     * @return 内置的 LocalDateTime 解析器
     */
    public static LocalDateTimeParser offeredLocalDateTimeParser() {
        return new LocalDateTimeParser();
    }

    /**
     * 内置的 SqlDate 解析器
     *
     * @return 内置的 SqlDate 解析器
     */
    public static SqlDateParser offeredSqlDateParser() {
        return new SqlDateParser();
    }

    /**
     * 内置的 Timestamp 解析器
     *
     * @return 内置的 Timestamp 解析器
     */
    public static TimestampParser offeredTimestampParser() {
        return new TimestampParser();
    }

    /**
     * 检查列注解，并构造相关组件实例
     *
     * @param ec    当前列信息
     * @param field 字段
     */
    public static void checkColumn(ExcelColumn ec, Field field) {
        // 过滤器
        Filter filter = field.getAnnotation(Filter.class);
        if (null != filter) {
            ec.setFilter(instantiate(filter.value()));
        }
        // 类型转换器
        Converter converter = field.getAnnotation(Converter.class);
        if (null != converter) {
            ec.setConverter(instantiate(converter.value()));
        }
        // 日期解析器
        Parser parser = field.getAnnotation(Parser.class);
        if (null != parser) {
            ec.setParser(instantiate(parser.value()));
        }
    }
}
