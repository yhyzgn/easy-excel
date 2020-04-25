package com.yhy.doc.excel.io;

import com.yhy.doc.excel.annotation.Excel;
import com.yhy.doc.excel.annotation.Ignored;
import com.yhy.doc.excel.annotation.Sorted;
import com.yhy.doc.excel.extra.ExcelTitle;
import com.yhy.doc.excel.internal.ExcelConstant;
import com.yhy.doc.excel.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:41
 * version: 1.0.0
 * desc   : Excel输出器
 */
public class ExcelWriter<T> {
    private final static String SUFFIX_XLS = ".xls";
    private final static String SUFFIX_XLSX = ".xlsx";
    private final static Map<String, String> MIME_TYPE = new HashMap<String, String>() {
        private static final long serialVersionUID = 5887513429547481187L;

        {
            put(SUFFIX_XLS, "application/vnd.ms-excel");
            put(SUFFIX_XLSX, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
    };

    private final OutputStream os;
    private HttpServletResponse response;
    private String filename;
    private String sheetName;
    private List<T> src;
    private Class<?> clazz;
    private String suffix;
    private boolean isBig;
    private Map<Field, ExcelTitle> titleMap = new TreeMap<>((o1, o2) -> o1.equals(o2) ? 0 : 1);

    public ExcelWriter(@NotNull File file) throws FileNotFoundException {
        this(new FileOutputStream(file));
    }

    public ExcelWriter(@NotNull HttpServletResponse response) throws Exception {
        this(response, null);
    }

    public ExcelWriter(@NotNull HttpServletResponse response, @Nullable String filename) throws Exception {
        response.reset();
        // 默认 xlsx 格式
        if (null == filename || "".equals(filename)) {
            filename = new SimpleDateFormat("yyyy-MM-dd HH_mm_ss.xlsx").format(Calendar.getInstance(Locale.getDefault()));
        }
        suffix = filename.substring(filename.lastIndexOf("."));
        if ("".equals(suffix) || !MIME_TYPE.containsKey(suffix)) {
            throw new IllegalStateException("unsupported file type: " + filename);
        }
        this.os = response.getOutputStream();
        this.response = response;
        this.filename = filename;
    }

    public ExcelWriter(@NotNull OutputStream os) {
        this.os = os;
        this.suffix = SUFFIX_XLS;
    }

    public ExcelWriter<T> x() {
        suffix = SUFFIX_XLSX;
        return this;
    }

    public ExcelWriter<T> big() {
        isBig = true;
        return x();
    }

    public void write(String sheetName, @NotNull T[] src) throws Exception {
        this.write(sheetName, Arrays.asList(src));
    }

    public void write(String sheetName, @NotNull Set<T> src) throws Exception {
        this.write(sheetName, new ArrayList<>(src));
    }

    public void write(String sheetName, @NotNull List<T> src) throws Exception {
        if (src.size() == 0) {
            return;
        }
        this.sheetName = sheetName;
        this.src = src;
        this.clazz = src.get(0).getClass();

        parseTitles();

        Workbook book = null;
        if (SUFFIX_XLS.equals(suffix)) {
            // xls
            book = writing();
        } else {
            if (isBig) {
                // xlsx && big data
                book = writingBig();
            } else {
                // xlsx
                book = writingX();
            }
        }

        if (null != response) {
            // 校验后缀，最终以 suffix 为准
            String temp = filename.substring(filename.lastIndexOf("."));
            if (!suffix.equals(temp)) {
                filename = filename.replace(temp, suffix);
            }
            response.setContentType(MIME_TYPE.get(suffix) + "; charset=utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + filename);
        }

        book.write(os);
        os.flush();

        release();
    }

    private void parseTitles() {
        List<Field> fields = new ArrayList<>(Arrays.asList(clazz.getDeclaredFields()));
        // 过滤字段，存储标题
        fields.stream().filter(field -> !field.isAnnotationPresent(Ignored.class) && !Modifier.isStatic(field.getModifiers())).sorted((f1, f2) -> {
            int s1 = fields.size(), s2 = fields.size();
            if (f1.isAnnotationPresent(Sorted.class)) {
                s1 = f1.getAnnotation(Sorted.class).value();
            }
            if (f2.isAnnotationPresent(Sorted.class)) {
                s2 = f2.getAnnotation(Sorted.class).value();
            }
            return s1 - s2;
        }).forEach(this::parseTitle);
    }

    private void parseTitle(Field field) {
        Excel excel = field.getAnnotation(Excel.class);
        String name = field.getName();
        if (null != excel) {
            if (!"".equals(excel.value())) {
                name = excel.value();
            } else if (!"".equals(excel.export())) {
                name = excel.export();
            }
        }

        // 将title添加到map中缓存
        ExcelTitle title = new ExcelTitle(name).setField(field);
        ExcelUtils.checkTitle(title, field);

        // 添加到缓存
        titleMap.put(field, title);
    }

    private void release() {
    }

    private Workbook writing() throws Exception {
        return wt(new HSSFWorkbook());
    }

    private Workbook writingX() throws Exception {
        return wt(new XSSFWorkbook());
    }

    private Workbook writingBig() throws Exception {
        return wt(new SXSSFWorkbook(1000));
    }

    private Workbook wt(Workbook book) throws Exception {
        Sheet sheet = book.getSheet(sheetName);
        if (null == sheet) {
            sheet = book.createSheet(sheetName);
        }

        sheet.setDefaultColumnWidth(ExcelConstant.COLUMN_SIZE);

        int rowIndex = sheet.getLastRowNum();
        if (rowIndex > 0) {
            rowIndex++;
        }

        // title
        writeTitle(sheet, rowIndex++);

        // data
        writeData(sheet, rowIndex);

        return book;
    }

    private void writeTitle(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        Cell cell;
        int index = 0;
        for (Map.Entry<Field, ExcelTitle> et : titleMap.entrySet()) {
            cell = row.createCell(index++);
            cell.setCellValue(et.getValue().getName());
        }
    }

    @SuppressWarnings("unchecked")
    private void writeData(Sheet sheet, int startRowIndex) throws Exception {
        T item;
        Row row;
        Cell cell;
        int titleIndex;
        Method getter;
        Object value;
        ExcelTitle title;

        for (int i = 0; i < src.size(); i++) {
            item = src.get(i);
            row = sheet.createRow(startRowIndex++);
            titleIndex = 0;
            for (Map.Entry<Field, ExcelTitle> et : titleMap.entrySet()) {
                title = et.getValue();
                cell = row.createCell(titleIndex++);
                getter = ExcelUtils.getter(et.getKey(), clazz);
                value = getter.invoke(item);
                if (null != title.getFilter()) {
                    value = title.getFilter().write(value);
                }
                if (null != title.getConverter()) {
                    value = title.getConverter().write(value);
                }
                if (null != title.getFormatter()) {
                    value = title.getFormatter().write(value);
                }
                cell.setCellValue(value.toString());
            }
        }
    }
}
