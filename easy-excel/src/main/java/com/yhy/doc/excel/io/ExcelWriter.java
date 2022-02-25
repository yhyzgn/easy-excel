package com.yhy.doc.excel.io;

import com.yhy.doc.excel.annotation.Font;
import com.yhy.doc.excel.annotation.*;
import com.yhy.doc.excel.extra.ExcelColumn;
import com.yhy.doc.excel.extra.Rect;
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
import java.net.URLEncoder;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;

/**
 * Excel 输出器
 * <p>
 * Created on 2019-09-09 12:41
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
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
    private Map<Field, Rect> autoMergeRowMap = new HashMap<>();

    /**
     * 写入Excel文件
     *
     * @param file 文件对象
     * @throws FileNotFoundException 文件异常
     */
    public ExcelWriter(@NotNull File file) throws FileNotFoundException {
        checkSuffix(file.getName());
        this.os = new FileOutputStream(file);
    }

    /**
     * 写入 HttpServletResponse 实现文件下载
     *
     * @param response HttpServletResponse
     * @throws Exception 可能出现的异常
     */
    public ExcelWriter(@NotNull HttpServletResponse response) throws Exception {
        this(response, ExcelUtils.defaultFilename());
    }

    /**
     * 写入 HttpServletResponse 实现文件下载
     *
     * @param response HttpServletResponse
     * @param filename 文件名
     * @throws Exception 可能出现的异常
     */
    public ExcelWriter(@NotNull HttpServletResponse response, @Nullable String filename) throws Exception {
        response.reset();
        if (StringUtils.isEmpty(filename)) {
            filename = ExcelUtils.defaultFilename();
        }
        if (!filename.contains(".")) {
            filename += SUFFIX_XLSX;
        }
        checkSuffix(filename);
        this.os = response.getOutputStream();
        this.response = response;
        this.filename = filename;
    }

    /**
     * 写入输出流中
     *
     * @param os 输出流
     */
    public ExcelWriter(@NotNull OutputStream os) {
        this.os = os;
        this.suffix = SUFFIX_XLS;
    }

    /**
     * 指定 xlsx 格式
     *
     * @return 当前实例
     */
    public ExcelWriter<T> x() {
        suffix = SUFFIX_XLSX;
        return this;
    }

    /**
     * 指定为大数据量写入
     *
     * @return 当前实例
     */
    public ExcelWriter<T> big() {
        isBig = true;
        return x();
    }

    /**
     * 指定数据源（数组），并开始写操作
     *
     * @param src 数据源
     * @throws Exception 可能出现的异常
     */
    public void write(@NotNull T[] src) throws Exception {
        this.write(ExcelUtils.defaultSheet(), Arrays.asList(src));
    }

    /**
     * 指定数据源（数组），并开始写操作
     *
     * @param sheetName 指定工作簿名称
     * @param src       数据源
     * @throws Exception 可能出现的异常
     */
    public void write(String sheetName, @NotNull T[] src) throws Exception {
        this.write(sheetName, Arrays.asList(src));
    }

    /**
     * 指定数据源（Set），并开始写操作
     *
     * @param src 数据源
     * @throws Exception 可能出现的异常
     */
    public void write(@NotNull Set<T> src) throws Exception {
        this.write(ExcelUtils.defaultSheet(), new ArrayList<>(src));
    }

    /**
     * 指定数据源（Set），并开始写操作
     *
     * @param sheetName 指定工作簿名称
     * @param src       数据源
     * @throws Exception 可能出现的异常
     */
    public void write(String sheetName, @NotNull Set<T> src) throws Exception {
        this.write(sheetName, new ArrayList<>(src));
    }

    /**
     * 指定数据源（List），并开始写操作
     *
     * @param src 数据源
     * @throws Exception 可能出现的异常
     */
    public void write(@NotNull List<T> src) throws Exception {
        this.write(ExcelUtils.defaultSheet(), src);
    }

    /**
     * 指定数据源（List），并开始写操作
     *
     * @param sheetName 指定工作簿名称
     * @param src       数据源
     * @throws Exception 可能出现的异常
     */
    public void write(String sheetName, @NotNull List<T> src) throws Exception {
        if (src.size() == 0) {
            return;
        }
        if (StringUtils.isEmpty(sheetName)) {
            sheetName = ExcelUtils.defaultSheet();
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
            response.setCharacterEncoding("utf-8");
            response.setContentType(String.format("%s; charset=utf-8", MIME_TYPE.get(suffix)));
            response.setHeader("Content-Disposition", "attachment; filename*=utf-8''" + URLEncoder.encode(filename, "utf-8"));
            response.addHeader("Pragma", "No-cache");
            response.addHeader("Cache-Control", "No-cache");
            response.setDateHeader("Expires", 0);
        }

        if (SUFFIX_XLS.equals(suffix)) {
            // xls
            book = workBook();
        } else {
            if (isBig) {
                // xlsx && big data
                book = workBookBig();
            } else {
                // xlsx
                book = workBookX();
            }
        }

        // 解析列信息
        parseColumns();

        // 开始写操作
        startWriter();

        if (null != book) {
            if (book.getNumberOfSheets() > 0) {
                book.setActiveSheet(0);
            }
            book.write(os);
        }
        os.flush();

        // 释放资源
        release();
    }

    /**
     * 检查文件名后缀
     *
     * @param filename 文件名
     * @throws IllegalStateException 后缀异常
     */
    private void checkSuffix(String filename) throws IllegalStateException {
        suffix = filename.substring(filename.lastIndexOf("."));
        if ("".equals(suffix) || !MIME_TYPE.containsKey(suffix)) {
            throw new IllegalStateException("unsupported file type: " + filename);
        }
    }

    /**
     * 解析列信息
     */
    private void parseColumns() {
        // 先解析表头样式
        if (clazz.isAnnotationPresent(Document.class)) {
            titleStyle = clazz.getAnnotation(Document.class).titleStyle();
            headerStyle = parseStyle(titleStyle);
        }

        List<Field> fields = new ArrayList<>(Arrays.asList(clazz.getDeclaredFields()));
        autoMergeRowMap.clear();
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

    /**
     * 解析字段信息
     *
     * @param field 字段
     */
    private void parseColumn(Field field) {
        field.setAccessible(true);

        // 提取自动合并单元格信息
        if (field.isAnnotationPresent(AutoMerge.class)) {
            Rect rect = null;
            Object last = null, current;
            for (int i = 0, start = 0, end = 0; i < this.src.size(); i++) {
                T src = this.src.get(i);
                try {
                    current = field.get(src);
                    if (null != current && null != last) {
                        if (current == last || current.equals(last)) {
                            end++;
                            last = current;
                            continue;
                        }
                    }
                    rect = new Rect(true, start, end, 0, 0);
                    start = end = i;
                    last = current;
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
                autoMergeRowMap.put(field, rect);
            }
        }

        Column column = field.getAnnotation(Column.class);
        String name = field.getName();
        String formula = null;
        if (null != column) {
            if (!"".equals(column.export())) {
                name = column.export();
            } else if (!"".equals(column.value())) {
                name = column.value();
            }
            formula = column.formula();
        }

        // 将column添加到map中缓存
        ExcelColumn ec = ExcelColumn.builder()
            .name(name)
            .field(field)
            .formula(formula)
            .mergeRect(autoMergeRowMap.get(field))
            .build();
        ExcelUtils.checkColumn(ec, field);

        // 解析style并缓存
        CellStyle cs = parseStyle(field, field.getAnnotation(Style.class), ec);
        if (null != cs) {
            styleMap.put(ec, cs);
        }
        // 添加到缓存
        columnMap.put(field, ec);
    }

    /**
     * 释放资源
     *
     * @throws Exception IO关闭异常
     */
    private void release() throws Exception {
        if (null != os) {
            os.close();
        }
        if (null != book) {
            book.close();
        }
        columnMap = null;
        autoMergeRowMap = null;
        styleMap = null;
    }

    /**
     * 创建工作簿，xls格式
     *
     * @return 工作簿
     */
    private Workbook workBook() {
        return new HSSFWorkbook();
    }

    /**
     * 创建工作簿，xlsx格式
     *
     * @return 工作簿
     */
    private Workbook workBookX() {
        return new XSSFWorkbook();
    }

    /**
     * 创建工作簿，xlsx格式，大数据量写入
     *
     * @return 工作簿
     */
    private Workbook workBookBig() {
        return new SXSSFWorkbook(1000);
    }

    /**
     * 开始写入操作
     *
     * @throws Exception 可能出现的异常
     */
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

    /**
     * 写入标题行
     *
     * @param sheet    指定工作簿
     * @param rowIndex 标题行索引
     */
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

    /**
     * 写入所有数据行（标题行除外）
     *
     * @param sheet         指定工作簿
     * @param startRowIndex 数据行开始索引
     * @throws Exception 可能出现的异常
     */
    @SuppressWarnings("unchecked")
    private void writeData(Sheet sheet, int startRowIndex) throws Exception {
        T item;
        Row row;
        Cell cell;
        int titleIndex;
        Method getter;
        Object value;
        ExcelColumn column;

        // TODO 合并单元格导出

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

    /**
     * 写入单元格
     *
     * @param cell     单元格
     * @param column   列信息
     * @param value    值
     * @param rowIndex 行索引
     */
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
            // 其他类型，先执行转换器，再写入
            if (null != column.getConverter()) {
                value = column.getConverter().write(value);
            }
            if (null != cs) {
                cell.setCellStyle(cs);
            }
            writeString(cell, value.toString());
        }
    }

    /**
     * 写入字符串类型数据
     *
     * @param cell  单元格
     * @param value 值
     */
    private void writeString(Cell cell, String value) {
        cell.setCellValue(value);
        // 检查是否是超链接，是则设置单元格为超链接格式
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

    /**
     * 数据格式化处理
     *
     * @param field      字段
     * @param defPattern 默认格式
     * @return 格式化方式索引
     */
    private short formatter(Field field, String defPattern) {
        Pattern pattern = field.getAnnotation(Pattern.class);
        return book.createDataFormat().getFormat(null != pattern ? pattern.value() : defPattern);
    }

    /**
     * 解析字段样式，主要用于各字段分别解析
     *
     * @param field  字段
     * @param style  样式定义
     * @param column 列信息
     * @return 单元格样式
     */
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

    /**
     * 解析样式，主要用于标题统一解析
     *
     * @param style 样式定义
     * @return 单元格样式
     */
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

    /**
     * 对齐格式
     *
     * @param cs    单元格样式
     * @param align 对齐格式定义
     */
    private void styleAlign(CellStyle cs, Align align) {
        if (null != align && align.enabled()) {
            cs.setAlignment(align.horizontal());
            cs.setVerticalAlignment(align.vertical());
            cs.setWrapText(align.wrap());
            cs.setIndention(align.indention());
            cs.setRotation(align.rotation());
        }
    }

    /**
     * 边框样式
     *
     * @param cs     单元格样式
     * @param border 边框样式定义
     */
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

    /**
     * 字体样式
     *
     * @param cs   单元格样式
     * @param font 字体样式定义
     */
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

    /**
     * 背景和底纹样式
     *
     * @param cs     单元格样式
     * @param ground 背景和底纹样式定义
     */
    private void styleGround(CellStyle cs, Ground ground) {
        if (null != ground && ground.enabled()) {
            cs.setFillBackgroundColor(ground.back().index);
            cs.setFillForegroundColor(ground.fore().index);
            cs.setFillPattern(ground.pattern());
        }
    }

    /**
     * 单元格尺寸格式
     *
     * @param size   尺寸定义
     * @param column 列信息
     */
    private void styleSize(Size size, ExcelColumn column) {
        if (null != size && size.enabled()) {
            column.setRowHeight(size.height());
        }
    }

    /**
     * 判断边框样式是否为某一边，ALL表示所有边
     *
     * @param border 边框样式定义
     * @param side   某一边
     * @return 是否为该边
     */
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
