package com.yhy.doc.excel.io;

import com.yhy.doc.excel.annotation.Excel;
import com.yhy.doc.excel.annotation.Ignored;
import com.yhy.doc.excel.annotation.Pattern;
import com.yhy.doc.excel.annotation.Sorted;
import com.yhy.doc.excel.extra.ExcelTitle;
import com.yhy.doc.excel.internal.EConstant;
import com.yhy.doc.excel.utils.ExcelUtils;
import com.yhy.doc.excel.utils.StringUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
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
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
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
    private Workbook book;
    private CreationHelper helper;
    private final Map<Field, ExcelTitle> titleMap = new TreeMap<>((o1, o2) -> o1.equals(o2) ? 0 : 1);

    public ExcelWriter(@NotNull File file) throws FileNotFoundException {
        parseSuffix(file.getName());
        this.os = new FileOutputStream(file);
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
        parseSuffix(filename);
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

        if (SUFFIX_XLS.equals(suffix)) {
            // xls
            writing();
        } else {
            if (isBig) {
                // xlsx && big data
                writingBig();
            } else {
                // xlsx
                writingX();
            }
        }

        if (null != response) {
            // 校验后缀，最终以 suffix 为准
            String temp = filename.substring(filename.lastIndexOf("."));
            if (!suffix.equals(temp)) {
                filename = filename.replace(temp, suffix);
            }
            response.setContentType(MIME_TYPE.get(suffix) + ";charset=utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + filename);
            response.setCharacterEncoding("UTF-8");
            response.addHeader("Pargam", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
        }

        if (null != book) {
            book.setActiveSheet(0);
            book.write(os);
        }
        os.flush();

        release();
    }

    private void parseSuffix(String filename) throws IllegalStateException {
        suffix = filename.substring(filename.lastIndexOf("."));
        if ("".equals(suffix) || !MIME_TYPE.containsKey(suffix)) {
            throw new IllegalStateException("unsupported file type: " + filename);
        }
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
        field.setAccessible(true);
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

    private void writing() throws Exception {
        wt(new HSSFWorkbook());
    }

    private void writingX() throws Exception {
        wt(new XSSFWorkbook());
    }

    private void writingBig() throws Exception {
        wt(new SXSSFWorkbook(1000));
    }

    private Workbook wt(Workbook bk) throws Exception {
        book = bk;
        helper = book.getCreationHelper();

        Sheet sheet = bk.getSheet(sheetName);
        if (null == sheet) {
            sheet = bk.createSheet(sheetName);
        }
        sheet.setDefaultColumnWidth(EConstant.COLUMN_SIZE);
        sheet.setVerticallyCenter(true);

        int rowIndex = sheet.getLastRowNum();

        // title
        writeTitle(sheet, ++rowIndex);

        // data
        writeData(sheet, ++rowIndex);

        return bk;
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
        ExcelTitle column;

        for (int i = 0; i < src.size(); i++) {
            item = src.get(i);
            row = sheet.createRow(startRowIndex++);
            titleIndex = 0;
            for (Map.Entry<Field, ExcelTitle> et : titleMap.entrySet()) {
                column = et.getValue();
                cell = row.createCell(titleIndex++);
                getter = ExcelUtils.getter(et.getKey(), clazz);
                value = getter.invoke(item);
                if (null != column.getFilter()) {
                    value = column.getFilter().write(value);
                }
                writeToCell(cell, column, value);
            }
        }
    }

    @SuppressWarnings("unchecked")
    private void writeToCell(Cell cell, ExcelTitle column, Object value) {
        if (null == value) {
            cell.setBlank();
            return;
        }

        Field field = column.getField();
        Class<?> type = field.getType();

        DataFormat format = book.createDataFormat();
        CellStyle style = style(field);
        Pattern pattern = field.getAnnotation(Pattern.class);

        if (type == String.class || type == CharSequence.class) {
            writeString(cell, value.toString());
        } else if (type == Integer.class || type == int.class) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : "#,#0"));
                cell.setCellStyle(style);
            }
            cell.setCellValue(Integer.parseInt(String.valueOf(value)));
        } else if (type == Float.class || type == float.class) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : "#,##0.00"));
                cell.setCellStyle(style);
            }
            cell.setCellValue(Float.parseFloat(String.valueOf(value)));
        } else if (type == Byte.class || type == byte.class) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : "#,#0"));
                cell.setCellStyle(style);
            }
            cell.setCellValue(Byte.parseByte(String.valueOf(value)));
        } else if (type == Boolean.class || type == boolean.class) {
            cell.setCellValue(Boolean.parseBoolean(String.valueOf(value)));
        } else if (type == Long.class || type == long.class) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : "#,#0"));
                cell.setCellStyle(style);
            }
            cell.setCellValue(Long.parseLong(String.valueOf(value)));
        } else if (type == Short.class || type == short.class) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : "#,#0"));
                cell.setCellStyle(style);
            }
            cell.setCellValue(Short.parseShort(String.valueOf(value)));
        } else if (type == Double.class || type == double.class) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : "#,##0.00"));
                cell.setCellStyle(style);
            }
            cell.setCellValue(Double.parseDouble(String.valueOf(value)));
        } else if ((type == Character.class || type == char.class) && value instanceof Character) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : "#,#0"));
                cell.setCellStyle(style);
            }
            cell.setCellValue((Character) value);
        } else if (type == Date.class && value instanceof Date) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : EConstant.PATTERN_DATE_TIME));
                cell.setCellStyle(style);
            }
            cell.setCellValue((Date) value);
        } else if (type == LocalDateTime.class && value instanceof LocalDateTime) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : EConstant.PATTERN_DATE_TIME));
                cell.setCellStyle(style);
            }
            cell.setCellValue((LocalDateTime) value);
        } else if (type == java.sql.Date.class && value instanceof java.sql.Date) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : EConstant.PATTERN_DATE_TIME));
                cell.setCellStyle(style);
            }
            cell.setCellValue((java.sql.Date) value);
        } else if (type == Timestamp.class && value instanceof Timestamp) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : EConstant.PATTERN_DATE_TIME));
                cell.setCellStyle(style);
            }
            cell.setCellValue((Timestamp) value);
        } else if (type == LocalDate.class && value instanceof LocalDate) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : EConstant.PATTERN_DATE));
                cell.setCellStyle(style);
            }
            cell.setCellValue((LocalDate) value);
        } else if (type == LocalTime.class && value instanceof LocalTime) {
            if (null != style) {
                style.setDataFormat(format.getFormat(null != pattern ? pattern.value() : EConstant.PATTERN_TIME));
                cell.setCellStyle(style);
            }
            cell.setCellValue(LocalDateTime.of(LocalDate.now(), (LocalTime) value));
        } else {
            if (null != column.getConverter()) {
                value = column.getConverter().write(value);
            }
            writeString(cell, value.toString());
        }
    }

    private void writeString(Cell cell, String value) {
        HyperlinkType type = StringUtils.hyperLinkType(value);
        if (type != HyperlinkType.NONE) {
            cell.setHyperlink(helper.createHyperlink(type));
        }
        cell.setCellValue(value);
    }

    private CellStyle style(Field field) {
        return book.createCellStyle();
//        return null;
    }
}
