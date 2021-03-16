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
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:14
 * version: 1.0.0
 * desc   : Excel对外工具类
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
     * 从Excel文件读取数据
     *
     * @param file  文件
     * @param clazz 映射的类
     * @param <T>   映射的类
     * @return 读取到的数据集
     */
    public static <T> List<T> read(File file, Class<T> clazz) {
        return read(file, ReaderConfig.deft(), clazz);
    }

    /**
     * 从Excel文件中读取数据
     *
     * @param file   文件
     * @param clazz  映射的类
     * @param config 读取配置
     * @param <T>    映射的类
     * @return 读取到的数据集
     */
    public static <T> List<T> read(File file, ReaderConfig config, Class<T> clazz) {
        try {
            return read(new FileInputStream(file), config, clazz);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    /**
     * 从输入流中读取数据
     * <p>
     * 如：FileInputStream(), MultipartFile.getInputStream() 等...
     *
     * @param is    输入流
     * @param clazz 映射的类
     * @param <T>   映射的类
     * @return 读取到的数据集
     */
    public static <T> List<T> read(InputStream is, Class<T> clazz) {
        return read(is, ReaderConfig.deft(), clazz);
    }

    /**
     * 从输入流中读取数据
     * <p>
     * 如：FileInputStream(), MultipartFile.getInputStream() 等...
     *
     * @param is     输入流
     * @param clazz  映射的类
     * @param config 读取配置
     * @param <T>    映射的类
     * @return 读取到的数据集
     */
    public static <T> List<T> read(InputStream is, ReaderConfig config, Class<T> clazz) {
        ExcelReader<T> reader = null;
        try {
            return new ExcelReader<T>(is, config).read(clazz);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    /**
     * 从ServletRequest中读取数据
     * <p>
     * 用于 Servlet 文件上传，仅限 binary 方式上传的文件读取
     *
     * @param request ServletRequest
     * @param clazz   映射的类
     * @param <T>     映射的类
     * @return 读取到的数据集
     */
    public static <T> List<T> read(ServletRequest request, Class<T> clazz) {
        return read(request, ReaderConfig.deft(), clazz);
    }

    /**
     * 从ServletRequest中读取数据
     * <p>
     * 用于 Servlet 文件上传，仅限 binary 方式上传的文件读取
     *
     * @param request ServletRequest
     * @param clazz   映射的类
     * @param config  读取配置
     * @param <T>     映射的类
     * @return 读取到的数据集
     */
    public static <T> List<T> read(ServletRequest request, ReaderConfig config, Class<T> clazz) {
        try {
            return read(request.getInputStream(), config, clazz);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    /**
     * 将数据写入Excel文件，xls格式
     *
     * @param file 文件对象
     * @param src  数据源，List
     * @param <T>  映射的类
     */
    public static <T> void write(File file, List<T> src) {
        write(file, src, defaultSheet());
    }

    /**
     * 将数据写入Excel文件，xls格式
     *
     * @param file      文件对象
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void write(File file, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(file).write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将数据写入Excel文件，xlsx格式
     *
     * @param file 文件对象
     * @param src  数据源
     * @param <T>  映射的类
     */
    public static <T> void writeX(File file, List<T> src) {
        writeX(file, src, defaultSheet());
    }

    /**
     * 将数据写入Excel文件，xlsx格式
     *
     * @param file      文件对象
     * @param src       数据源
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void writeX(File file, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(file).x().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将数据写入Excel文件，xlsx格式，大数据量写入
     *
     * @param file 文件对象
     * @param src  数据源
     * @param <T>  映射的类
     */
    public static <T> void writeBig(File file, List<T> src) {
        writeBig(file, src, defaultSheet());
    }

    /**
     * 将数据写入Excel文件，xlsx格式，大数据量写入
     *
     * @param file      文件对象
     * @param src       数据源
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void writeBig(File file, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(file).x().big().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将数据写入输出流中，xls格式
     *
     * @param os  输出流
     * @param src 数据源，List
     * @param <T> 映射的类
     */
    public static <T> void write(OutputStream os, List<T> src) {
        write(os, src, defaultSheet());
    }

    /**
     * 将数据写入输出流中，xls格式
     *
     * @param os        输出流
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void write(OutputStream os, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(os).write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将数据写入输出流中，xlsx格式
     *
     * @param os  输出流
     * @param src 数据源，List
     * @param <T> 映射的类
     */
    public static <T> void writeX(OutputStream os, List<T> src) {
        writeX(os, src, defaultSheet());
    }

    /**
     * 将数据写入输出流中，xlsx格式
     *
     * @param os        输出流
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void writeX(OutputStream os, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(os).x().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将数据写入输出流中，xlsx格式，大数据量写入
     *
     * @param os  输出流
     * @param src 数据源，List
     * @param <T> 映射的类
     */
    public static <T> void writeBig(OutputStream os, List<T> src) {
        writeBig(os, src, defaultSheet());
    }

    /**
     * 将数据写入输出流中，xlsx格式，大数据量写入
     *
     * @param os        输出流
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void writeBig(OutputStream os, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(os).x().big().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xls格式
     *
     * @param response HttpServletResponse
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    public static <T> void write(HttpServletResponse response, List<T> src) {
        write(response, src, defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xls格式
     *
     * @param response  HttpServletResponse
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void write(HttpServletResponse response, List<T> src, String sheetName) {
        write(response, defaultFilename(), src, sheetName);
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式
     *
     * @param response HttpServletResponse
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    public static <T> void writeX(HttpServletResponse response, List<T> src) {
        writeX(response, src, defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式
     *
     * @param response  HttpServletResponse
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void writeX(HttpServletResponse response, List<T> src, String sheetName) {
        writeX(response, defaultFilename(), src, sheetName);
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式，大数据量写入
     *
     * @param response HttpServletResponse
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    public static <T> void writeBig(HttpServletResponse response, List<T> src) {
        writeBig(response, src, defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式，大数据量写入
     *
     * @param response  HttpServletResponse
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void writeBig(HttpServletResponse response, List<T> src, String sheetName) {
        writeBig(response, defaultFilename(), src, sheetName);
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xls格式
     *
     * @param response HttpServletResponse
     * @param filename 下载时的文件名
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    public static <T> void write(HttpServletResponse response, String filename, List<T> src) {
        write(response, filename, src, defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xls格式
     *
     * @param response  HttpServletResponse
     * @param filename  下载时的文件名
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void write(HttpServletResponse response, String filename, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(response, filename).write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式
     *
     * @param response HttpServletResponse
     * @param filename 下载时的文件名
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    public static <T> void writeX(HttpServletResponse response, String filename, List<T> src) {
        writeX(response, filename, src, defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式
     *
     * @param response  HttpServletResponse
     * @param filename  下载时的文件名
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void writeX(HttpServletResponse response, String filename, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(response, filename).x().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式，大数据量写入
     *
     * @param response HttpServletResponse
     * @param filename 下载时的文件名
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    public static <T> void writeBig(HttpServletResponse response, String filename, List<T> src) {
        writeBig(response, filename, src, defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式，大数据量写入
     *
     * @param response  HttpServletResponse
     * @param filename  下载时的文件名
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    public static <T> void writeBig(HttpServletResponse response, String filename, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(response, filename).x().big().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
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
                return new Rect(true, firstRow, lastRow, firstColumn, lastColumn);
            }
        }
        return new Rect(false, row, row, column, column);
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
     * @param column 当前列信息
     * @param field  字段
     */
    public static void checkColumn(ExcelColumn column, Field field) {
        // 过滤器
        Filter filter = field.getAnnotation(Filter.class);
        if (null != filter && filter.value() != EFilter.class) {
            column.setFilter(instantiate(filter.value()));
        }
        // 类型转换器
        Converter converter = field.getAnnotation(Converter.class);
        if (null != converter && converter.value() != EConverter.class) {
            column.setConverter(instantiate(converter.value()));
        }
        // 日期解析器
        Parser parser = field.getAnnotation(Parser.class);
        if (null != parser && parser.value() != EDateParser.class) {
            column.setParser(instantiate(parser.value()));
        }
    }
}
