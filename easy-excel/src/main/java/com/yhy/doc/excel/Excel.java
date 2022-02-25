package com.yhy.doc.excel;

import com.yhy.doc.excel.extra.ReaderConfig;
import com.yhy.doc.excel.io.ExcelReader;
import com.yhy.doc.excel.io.ExcelWriter;
import com.yhy.doc.excel.utils.ExcelUtils;

import javax.servlet.ServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Collections;
import java.util.List;

/**
 * Excel 操作类
 * <p>
 * Created on 2022-02-25 9:43
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
public interface Excel {

    /**
     * 从Excel文件读取数据
     *
     * @param file  文件
     * @param clazz 映射的类
     * @param <T>   映射的类
     * @return 读取到的数据集
     */
    static <T> List<T> read(File file, Class<T> clazz) {
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
    static <T> List<T> read(File file, ReaderConfig config, Class<T> clazz) {
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
    static <T> List<T> read(InputStream is, Class<T> clazz) {
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
    static <T> List<T> read(InputStream is, ReaderConfig config, Class<T> clazz) {
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
    static <T> List<T> read(ServletRequest request, Class<T> clazz) {
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
    static <T> List<T> read(ServletRequest request, ReaderConfig config, Class<T> clazz) {
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
    static <T> void write(File file, List<T> src) {
        write(file, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入Excel文件，xls格式
     *
     * @param file      文件对象
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void write(File file, List<T> src, String sheetName) {
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
    static <T> void writeX(File file, List<T> src) {
        writeX(file, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入Excel文件，xlsx格式
     *
     * @param file      文件对象
     * @param src       数据源
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void writeX(File file, List<T> src, String sheetName) {
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
    static <T> void writeBig(File file, List<T> src) {
        writeBig(file, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入Excel文件，xlsx格式，大数据量写入
     *
     * @param file      文件对象
     * @param src       数据源
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void writeBig(File file, List<T> src, String sheetName) {
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
    static <T> void write(OutputStream os, List<T> src) {
        write(os, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入输出流中，xls格式
     *
     * @param os        输出流
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void write(OutputStream os, List<T> src, String sheetName) {
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
    static <T> void writeX(OutputStream os, List<T> src) {
        writeX(os, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入输出流中，xlsx格式
     *
     * @param os        输出流
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void writeX(OutputStream os, List<T> src, String sheetName) {
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
    static <T> void writeBig(OutputStream os, List<T> src) {
        writeBig(os, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入输出流中，xlsx格式，大数据量写入
     *
     * @param os        输出流
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void writeBig(OutputStream os, List<T> src, String sheetName) {
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
    static <T> void write(HttpServletResponse response, List<T> src) {
        write(response, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xls格式
     *
     * @param response  HttpServletResponse
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void write(HttpServletResponse response, List<T> src, String sheetName) {
        write(response, ExcelUtils.defaultFilename(), src, sheetName);
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式
     *
     * @param response HttpServletResponse
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    static <T> void writeX(HttpServletResponse response, List<T> src) {
        writeX(response, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式
     *
     * @param response  HttpServletResponse
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void writeX(HttpServletResponse response, List<T> src, String sheetName) {
        writeX(response, ExcelUtils.defaultFilename(), src, sheetName);
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式，大数据量写入
     *
     * @param response HttpServletResponse
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    static <T> void writeBig(HttpServletResponse response, List<T> src) {
        writeBig(response, src, ExcelUtils.defaultSheet());
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xlsx格式，大数据量写入
     *
     * @param response  HttpServletResponse
     * @param src       数据源，List
     * @param sheetName 工作簿名称
     * @param <T>       映射的类
     */
    static <T> void writeBig(HttpServletResponse response, List<T> src, String sheetName) {
        writeBig(response, ExcelUtils.defaultFilename(), src, sheetName);
    }

    /**
     * 将数据写入 HttpServletResponse 中，实现文件下载，xls格式
     *
     * @param response HttpServletResponse
     * @param filename 下载时的文件名
     * @param src      数据源，List
     * @param <T>      映射的类
     */
    static <T> void write(HttpServletResponse response, String filename, List<T> src) {
        write(response, filename, src, ExcelUtils.defaultSheet());
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
    static <T> void write(HttpServletResponse response, String filename, List<T> src, String sheetName) {
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
    static <T> void writeX(HttpServletResponse response, String filename, List<T> src) {
        writeX(response, filename, src, ExcelUtils.defaultSheet());
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
    static <T> void writeX(HttpServletResponse response, String filename, List<T> src, String sheetName) {
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
    static <T> void writeBig(HttpServletResponse response, String filename, List<T> src) {
        writeBig(response, filename, src, ExcelUtils.defaultSheet());
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
    static <T> void writeBig(HttpServletResponse response, String filename, List<T> src, String sheetName) {
        try {
            new ExcelWriter<T>(response, filename).x().big().write(sheetName, src);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
