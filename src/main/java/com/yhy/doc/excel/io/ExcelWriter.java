package com.yhy.doc.excel.io;

import com.yhy.doc.excel.annotation.Font;
import com.yhy.doc.excel.annotation.*;
import com.yhy.doc.excel.extra.ExcelColumn;
import com.yhy.doc.excel.internal.EBorderSide;
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
    private Style titleStyle;
    private CellStyle headerStyle;
    private Map<Field, ExcelColumn> columnMap = new TreeMap<>((o1, o2) -> o1.equals(o2) ? 0 : 1);
    private Map<ExcelColumn, CellStyle> styleMap = new HashMap<>();

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

        if (null != response) {
            // 校验后缀，最终以 suffix 为准
            String temp = filename.substring(filename.lastIndexOf("."));
            if (!suffix.equals(temp)) {
                filename = filename.replace(temp, suffix);
            }
            response.setContentType(String.format("%s; charset=utf-8", MIME_TYPE.get(suffix)));
            response.setHeader("Content-Disposition", "attachment; filename=" + filename);
            response.setCharacterEncoding("UTF-8");
            response.addHeader("Pragma", "public");
            response.addHeader("Cache-Control", "public");
        }

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

        parseColumns();

        startWriter();

        if (null != book) {
            if (book.getNumberOfSheets() > 0) {
                book.setActiveSheet(0);
            }
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

    private void parseColumns() {
        // 先解析表头样式
        if (clazz.isAnnotationPresent(Document.class)) {
            titleStyle = clazz.getAnnotation(Document.class).titleStyle();
            headerStyle = parseStyle(titleStyle);
        }

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
        }).forEach(this::parseColumn);
    }

    private void parseColumn(Field field) {
        field.setAccessible(true);
        Excel excel = field.getAnnotation(Excel.class);
        String name = field.getName();
        String formula = null;
        if (null != excel) {
            if (!"".equals(excel.value())) {
                name = excel.value();
            } else if (!"".equals(excel.export())) {
                name = excel.export();
            }
            formula = excel.formula();
        }

        // 将column添加到map中缓存
        ExcelColumn column = new ExcelColumn(name).setField(field).setFormula(formula);
        ExcelUtils.checkColumn(column, field);

        // 解析style并缓存
        CellStyle cs = parseStyle(field, field.getAnnotation(Style.class), column);
        if (null != cs) {
            styleMap.put(column, cs);
        }
        // 添加到缓存
        columnMap.put(field, column);
    }

    private void release() throws Exception {
        if (null != os) {
            os.close();
        }
        if (null != book) {
            book.close();
        }
        columnMap = null;
        styleMap = null;
    }

    private Workbook writing() throws Exception {
        return new HSSFWorkbook();
    }

    private Workbook writingX() throws Exception {
        return new XSSFWorkbook();
    }

    private Workbook writingBig() throws Exception {
        return new SXSSFWorkbook(1000);
    }

    private void startWriter() throws Exception {
        helper = book.getCreationHelper();

        Sheet sheet = book.getSheet(sheetName);
        if (null == sheet) {
            sheet = book.createSheet(sheetName);
        }
        sheet.setDefaultColumnWidth(EConstant.COLUMN_WIDTH);
        if (null != titleStyle) {
            // 自定义列宽（此处只设置列宽，行高不能在此统一设置）
            sheet.setDefaultColumnWidth(titleStyle.size().width());
        }
        sheet.setVerticallyCenter(true);

        int rowIndex = sheet.getLastRowNum();

        // title
        writeTitle(sheet, ++rowIndex);

        // data
        writeData(sheet, ++rowIndex);
    }

    private void writeTitle(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        if (null != titleStyle) {
            // 自定义行高（标题）
            row.setHeightInPoints(titleStyle.size().height());
        }
        Cell cell;
        int index = 0;
        for (Map.Entry<Field, ExcelColumn> et : columnMap.entrySet()) {
            cell = row.createCell(index++);
            if (null != headerStyle) {
                cell.setCellStyle(headerStyle);
            }
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
        ExcelColumn column;

        for (T t : src) {
            item = t;
            row = sheet.createRow(startRowIndex++);
            titleIndex = 0;
            for (Map.Entry<Field, ExcelColumn> et : columnMap.entrySet()) {
                column = et.getValue();
                cell = row.createCell(titleIndex++);
                // 执行字段对应的getter方法
                value = ExcelUtils.invokeGetter(item, et.getKey());
                if (null != column.getFilter()) {
                    value = column.getFilter().write(value);
                }
                // 设置行高
                if (column.getRowHeight() > 0) {
                    row.setHeightInPoints(column.getRowHeight());
                }
                writeToCell(cell, column, value, startRowIndex);
            }
        }
    }

    @SuppressWarnings("unchecked")
    private void writeToCell(Cell cell, ExcelColumn column, Object value, int rowIndex) {
        Field field = column.getField();
        Class<?> type = field.getType();
        CellStyle cs = styleMap.get(column);

        if (StringUtils.isNotEmpty(column.getFormula())) {
            // 把占位符换成行号
            String formula = column.getFormula().replaceAll("\\{}", String.valueOf(rowIndex));
            // 函数表达式
            if (null != cs) {
                cell.setCellStyle(cs);
            }
            cell.setCellFormula(formula);
            return;
        }

        if (null == value) {
            if (null != cs) {
                cell.setCellStyle(cs);
            }
            cell.setBlank();
            return;
        }

        // 其他数值类型
        if (type == String.class || type == CharSequence.class) {
            if (null != cs) {
                cell.setCellStyle(cs);
            }
            writeString(cell, value.toString());
        } else if (type == Integer.class || type == int.class) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, "#,#0"));
                cell.setCellStyle(cs);
            }
            cell.setCellValue(Integer.parseInt(String.valueOf(value)));
        } else if (type == Float.class || type == float.class) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, "#,#0"));
                cell.setCellStyle(cs);
            }
            cell.setCellValue(Float.parseFloat(String.valueOf(value)));
        } else if (type == Byte.class || type == byte.class) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, "#,#0"));
                cell.setCellStyle(cs);
            }
            cell.setCellValue(Byte.parseByte(String.valueOf(value)));
        } else if (type == Boolean.class || type == boolean.class) {
            if (null != cs) {
                cell.setCellStyle(cs);
            }
            cell.setCellValue(Boolean.parseBoolean(String.valueOf(value)));
        } else if (type == Long.class || type == long.class) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, "#,#0"));
                cell.setCellStyle(cs);
            }
            cell.setCellValue(Long.parseLong(String.valueOf(value)));
        } else if (type == Short.class || type == short.class) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, "#,#0"));
                cell.setCellStyle(cs);
            }
            cell.setCellValue(Short.parseShort(String.valueOf(value)));
        } else if (type == Double.class || type == double.class) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, "#,##0.00"));
                cell.setCellStyle(cs);
            }
            cell.setCellValue(Double.parseDouble(String.valueOf(value)));
        } else if ((type == Character.class || type == char.class) && value instanceof Character) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, "#,#0"));
                cell.setCellStyle(cs);
            }
            cell.setCellValue((Character) value);
        } else if (type == Date.class && value instanceof Date) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, EConstant.PATTERN_DATE_TIME));
                cell.setCellStyle(cs);
            }
            cell.setCellValue((Date) value);
        } else if (type == LocalDateTime.class && value instanceof LocalDateTime) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, EConstant.PATTERN_DATE_TIME));
                cell.setCellStyle(cs);
            }
            cell.setCellValue((LocalDateTime) value);
        } else if (type == java.sql.Date.class && value instanceof java.sql.Date) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, EConstant.PATTERN_DATE_TIME));
                cell.setCellStyle(cs);
            }
            cell.setCellValue((java.sql.Date) value);
        } else if (type == Timestamp.class && value instanceof Timestamp) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, EConstant.PATTERN_DATE_TIME));
                cell.setCellStyle(cs);
            }
            cell.setCellValue((Timestamp) value);
        } else if (type == LocalDate.class && value instanceof LocalDate) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, EConstant.PATTERN_DATE));
                cell.setCellStyle(cs);
            }
            cell.setCellValue((LocalDate) value);
        } else if (type == LocalTime.class && value instanceof LocalTime) {
            if (null != cs) {
                cs.setDataFormat(formatter(field, EConstant.PATTERN_TIME));
                cell.setCellStyle(cs);
            }
            cell.setCellValue(LocalDateTime.of(LocalDate.now(), (LocalTime) value));
        } else {
            if (null != column.getConverter()) {
                value = column.getConverter().write(value);
            }
            if (null != cs) {
                cell.setCellStyle(cs);
            }
            writeString(cell, value.toString());
        }
    }

    private void writeString(Cell cell, String value) {
        cell.setCellValue(value);
        HyperlinkType type = StringUtils.hyperLinkType(value);
        if (type != HyperlinkType.NONE) {
            if (type == HyperlinkType.EMAIL && !value.startsWith("mailto:")) {
                value = "mailto:" + value;
            }
            Hyperlink link = helper.createHyperlink(type);
            link.setAddress(value);
            cell.setHyperlink(link);
        }
    }

    private short formatter(Field field, String defPattern) {
        Pattern pattern = field.getAnnotation(Pattern.class);
        return book.createDataFormat().getFormat(null != pattern ? pattern.value() : defPattern);
    }

    private CellStyle parseStyle(Field field, Style style, ExcelColumn column) {
        if (null != style) {
            return parseStyle(style);
        }
        CellStyle cs = book.createCellStyle();
        // 单元格对齐方式
        styleAlign(cs, field.getAnnotation(Align.class));
        // 边框样式
        styleBorder(cs, field.getAnnotation(Border.class));
        // 字体样式
        styleFont(cs, field.getAnnotation(Font.class));
        // 背景和纹理
        styleGround(cs, field.getAnnotation(Ground.class));
        // 尺寸
        styleSize(field.getAnnotation(Size.class), column);
        return cs;
    }

    private CellStyle parseStyle(Style style) {
        CellStyle cs = book.createCellStyle();
        if (null != style) {
            // 单元格对齐方式
            styleAlign(cs, style.align());
            // 边框样式
            styleBorder(cs, style.border());
            // 字体样式
            styleFont(cs, style.font());
            // 背景和纹理
            styleGround(cs, style.ground());
        }
        return cs;
    }

    private void styleAlign(CellStyle cs, Align align) {
        if (null != align && align.enabled()) {
            cs.setAlignment(align.horizontal());
            cs.setVerticalAlignment(align.vertical());
            cs.setWrapText(align.wrap());
            cs.setIndention(align.indention());
            cs.setRotation(align.rotation());
        }
    }

    private void styleBorder(CellStyle cs, Border border) {
        if (null != border && border.enabled()) {
            if (borderSideIs(border, EBorderSide.LEFT)) {
                // 左
                cs.setBorderLeft(border.style());
                cs.setLeftBorderColor(border.color().index);
            }
            if (borderSideIs(border, EBorderSide.TOP)) {
                // 上
                cs.setBorderTop(border.style());
                cs.setTopBorderColor(border.color().index);
            }
            if (borderSideIs(border, EBorderSide.RIGHT)) {
                // 右
                cs.setBorderRight(border.style());
                cs.setRightBorderColor(border.color().index);
            }
            if (borderSideIs(border, EBorderSide.BOTTOM)) {
                // 下
                cs.setBorderBottom(border.style());
                cs.setBottomBorderColor(border.color().index);
            }
        }
    }

    private void styleFont(CellStyle cs, Font font) {
        if (null != font && font.enabled()) {
            org.apache.poi.ss.usermodel.Font ft = book.createFont();
            ft.setFontName(font.name());
            ft.setFontHeightInPoints(font.size());
            ft.setColor(font.color().index);
            ft.setUnderline(font.underline());
            ft.setTypeOffset(font.typeOffset());
            ft.setStrikeout(font.delete());
            cs.setFont(ft);
        }
    }

    private void styleGround(CellStyle cs, Ground ground) {
        if (null != ground && ground.enabled()) {
            cs.setFillBackgroundColor(ground.back().index);
            cs.setFillForegroundColor(ground.fore().index);
            cs.setFillPattern(ground.pattern());
        }
    }

    private void styleSize(Size size, ExcelColumn column) {
        if (null != size && size.enabled()) {
            column.setRowHeight(size.height());
        }
    }

    private boolean borderSideIs(Border border, EBorderSide side) {
        EBorderSide[] sides = border.sides();
        for (EBorderSide sd : sides) {
            if (sd == EBorderSide.ALL || sd == side) {
                return true;
            }
        }
        return false;
    }
}
