package com.yhy.doc.excel.io;

import com.yhy.doc.excel.annotation.Column;
import com.yhy.doc.excel.extra.CosineSimilarity;
import com.yhy.doc.excel.extra.ExcelColumn;
import com.yhy.doc.excel.extra.ReaderConfig;
import com.yhy.doc.excel.extra.Rect;
import com.yhy.doc.excel.utils.ExcelUtils;
import com.yhy.doc.excel.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.jetbrains.annotations.NotNull;

import javax.servlet.ServletRequest;
import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;

/**
 * Excel 读取器
 * <p>
 * Created on 2019-09-09 12:41
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Slf4j
public class ExcelReader<T> {
    private final InputStream is;
    private final ReaderConfig config;
    private final Workbook workbook;
    private Map<Integer, String> columnMap;
    private Map<Field, Integer> fieldIndexMap;
    private List<Map<Integer, Object>> valueList;
    private Map<Integer, ExcelColumn> excelColumnMap;
    private Class<T> clazz;
    private Sheet sheet;
    private Constructor<T> constructor;
    private int sheetIndex;
    private List<T> resultList;

    /**
     * 从Excel文件中读取数据
     *
     * @param file   文件
     * @param config 读取配置
     * @throws FileNotFoundException 文件异常
     */
    public ExcelReader(File file, ReaderConfig config) throws FileNotFoundException {
        this(new FileInputStream(file), config);
    }

    /**
     * 从 ServletRequest 中读取数据，实现文件上传
     *
     * @param request ServletRequest
     * @param config  读取配置
     * @throws IOException 可能出现的异常
     */
    public ExcelReader(ServletRequest request, ReaderConfig config) throws IOException {
        this(request.getInputStream(), config);
    }

    /**
     * 从输入流中读取数据
     *
     * @param is     输入流对象
     * @param config 读取配置
     */
    public ExcelReader(InputStream is, ReaderConfig config) {
        this.is = is;
        this.config = config;
        this.workbook = getWorkbook();
        this.fieldIndexMap = new HashMap<>();
        this.columnMap = new HashMap<>();
        this.valueList = new ArrayList<>();
        this.excelColumnMap = new HashMap<>();
        validate();
    }

    /**
     * 读取操作
     *
     * @param clazz 映射的类
     * @return 读取到的数据集
     * @throws IOException 可能出现的异常
     */
    public List<T> read(@NotNull Class<T> clazz) throws IOException {
        try {
            constructor = clazz.getConstructor();
        } catch (NoSuchMethodException e) {
            throw new IllegalArgumentException("Your model class '" + clazz.getName() + "' must contains a constructor without any argument, but not found.");
        }
        this.clazz = clazz;
        this.sheetIndex = Math.min(config.getSheetIndex(), getSheetCount() - 1);
        // 开始读取
        reading();
        // 释放资源
        List<T> temp = new ArrayList<>(resultList);
        release();
        return temp;
    }

    /**
     * 释放资源
     *
     * @throws IOException 可能出现的异常
     */
    public void release() throws IOException {
        if (null != is) {
            is.close();
        }
        if (null != workbook) {
            workbook.close();
        }
        columnMap = null;
        fieldIndexMap = null;
        valueList = null;
        excelColumnMap = null;
        sheet = null;
        resultList = null;
    }

    /**
     * 开始读取操作
     */
    private void reading() {
        sheet = workbook.getSheetAt(sheetIndex);
        // sheet.getPhysicalNumberOfRows() 方法获取到的行数会自动忽略合并的单元格
        int lastRowIndex = config.getRowEndIndex() > -1 ? config.getRowEndIndex() : sheet.getLastRowNum();
        if (lastRowIndex == -1) {
            // 没有任何数据
            return;
        }
        // 开始行的索引，不设置的话，默认从标题的下一行开始
        int firstRowIndex = config.getRowStartIndex() > 0 ? config.getRowStartIndex() : config.getTitleIndex() + 1;
        int rows = lastRowIndex - firstRowIndex + 1;

        // 读取标题
        readColumn();

        // 读取其他行
        readRows(firstRowIndex, rows);

        parse();
    }

    /**
     * 读取所有行（数据行，不包括标题）
     *
     * @param rowStart 行开始索引
     * @param rows     总行数（不包括标题）
     */
    private void readRows(int rowStart, int rows) {
        Row row;
        Cell cell;
        Rect rect;
        Object value;
        Map<Integer, Object> valuesOfRow;
        int columnStart = config.getCellStartIndex();
        for (int i = rowStart; i < rows + rowStart; i++) {
            valuesOfRow = new HashMap<>();
            row = sheet.getRow(i);
            if (null != row) {
                // row.getPhysicalNumberOfCells()  获取有记录的列数，即：最后有数据的列是第n列，前面有m列是空列没数据，则返回n-m；
                int cells = config.getCellEndIndex() > -1 ? config.getCellEndIndex() : row.getLastCellNum();
                if (cells == -1) {
                    // 该行没有任何数据
                    continue;
                }
                for (int j = columnStart; j < cells + columnStart; j++) {
                    cell = row.getCell(j);
                    if (null != cell) {
                        rect = ExcelUtils.merged(sheet, i, j);
                        if (rect.isMerged() && rect.getColumnStart() <= rect.getColumnEnd()) {
                            // 已合并的单元格
                            Row tempRow = sheet.getRow(rect.getRowStart());
                            Cell tempCell = tempRow.getCell(rect.getColumnStart());
                            value = getValueOfCell(tempCell, false);
                        } else {
                            // 其他，表格值
                            value = getValueOfCell(cell, false);
                        }
                        valuesOfRow.put(j, value);
                    }
                }
                valueList.add(valuesOfRow);
            }
        }
    }

    /**
     * 读取列ø
     */
    private void readColumn() {
        Row row = sheet.getRow(config.getTitleIndex());
        // 列的开始索引
        if (null != row) {
            Cell cell;
            Object value;
            int start = config.getCellStartIndex();
            int cells = row.getLastCellNum();
            if (cells == -1) {
                // 该行没有任何数据
                return;
            }
            List<String> titles = new ArrayList<>();
            for (int j = start; j < cells; j++) {
                cell = row.getCell(j);
                if (null != cell) {
                    value = getValueOfCell(cell, true);
                    titles.add(String.valueOf(value));
                }
            }

            List<Field> fields = new ArrayList<>(Arrays.asList(clazz.getDeclaredFields()));
            Column column;
            String title, name, like;
            for (int i = 0; i < titles.size(); i++) {
                title = resolveWrap(titles.get(i));
                for (Field item : fields) {
                    if (item.isAnnotationPresent(Column.class)) {
                        column = item.getAnnotation(Column.class);
                        name = resolveWrap(column.value());
                        if (column.insensitive()) {
                            // 忽略大小写，全部转为小写，以便匹配
                            title = title.toLowerCase();
                            name = name.toLowerCase();
                        }

                        // 先进行name完全匹配
                        if (name.equals(title)) {
                            columnMap.put(i, title);
                            fieldIndexMap.put(item, i);
                            break;
                        }

                        // 未匹配到正确的标题，进行模糊匹配
                        like = column.like().trim().replaceAll("%+", ".*?");
                        if (title.matches(like)) {
                            columnMap.put(i, title);
                            fieldIndexMap.put(item, i);
                            break;
                        }

                        // 如果还是未匹配到并且开启了智能匹配，就进行智能匹配
                        // 根据相似度容差查询列索引
                        if (column.intelligent()) {
                            // 求得相似度
                            double similarity = CosineSimilarity.getSimilarity(name, title);
                            if (similarity >= 1.0D - column.tolerance()) {
                                // 相似度在容差范围以内，表示匹配成功
                                columnMap.put(i, title);
                                fieldIndexMap.put(item, i);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 获取数据类型
     *
     * @param cell    单元格对象
     * @param isTitle 是否是标题
     * @return 单元格值
     */
    private Object getValueOfCell(Cell cell, boolean isTitle) {
        //判断是否为null或空串
        if (null == cell || "".equals(cell.toString().trim())) {
            return null;
        }
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        CellType type = evaluator.evaluate(cell).getCellType();
        Object value;
        switch (type) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BOOLEAN:
                value = isTitle ? String.valueOf(cell.getBooleanCellValue()) : cell.getBooleanCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // 日期时间，转换为毫秒
                    value = isTitle ? String.valueOf(cell.getDateCellValue().getTime()) : cell.getDateCellValue();
                } else {
                    value = new DecimalFormat("#.#########").format(cell.getNumericCellValue());
                }
                break;
            case BLANK:
                value = null;
                break;
            case FORMULA:
                value = evaluator.evaluate(cell).getStringValue();
                break;
            default:
                try {
                    value = cell.getStringCellValue();
                } catch (Exception e) {
                    e.printStackTrace();
                    value = null;
                }
                break;
        }
        if (value instanceof String && ((String) value).trim().isEmpty()) {
            value = null;
        }
        return value;
    }

    /**
     * 解析读取到的数据
     */
    private void parse() {
        if (columnMap.isEmpty()) {
            throw new IllegalStateException("Can not read columns of excel file.");
        }

        parseColumns();

        parseData();
    }

    /**
     * 解析读取到的数据
     */
    @SuppressWarnings("unchecked")
    private void parseData() {
        if (!valueList.isEmpty()) {
            resultList = new ArrayList<>();
            valueList.forEach(item -> {
                try {
                    T data = constructor.newInstance();
                    Integer index;
                    Object value;
                    ExcelColumn column;
                    for (Map.Entry<Integer, Object> et : item.entrySet()) {
                        index = et.getKey();
                        value = et.getValue();
                        column = excelColumnMap.get(index);

                        if (null == column) {
                            continue;
                        }
                        if (null == value && !column.isNullable()) {
                            return;
                        }

                        // 自动处理换行符
                        if (column.isWrap()) {
                            value = resolveWrap(String.valueOf(value));
                        }

                        // 先执行过滤器
                        if (null != column.getFilter()) {
                            value = column.getFilter().read(value);
                        }

                        // 执行转换器，格式化一些值得转换，比如枚举
                        if (null != column.getConverter()) {
                            value = column.getConverter().read(value);
                        }

                        // 类型转换
                        value = caseType(value, column);

                        // 如果value为null，就不需要设置啦~
                        if (null != value) {
                            // 执行字段对应的setter方法
                            ExcelUtils.invokeSetter(data, column.getField(), value);
                        }
                    }
                    resultList.add(data);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });
        }
    }

    /**
     * 解析所有列
     */
    private void parseColumns() {
        List<Field> fields = new ArrayList<>(Arrays.asList(clazz.getDeclaredFields()));
        // 将标题信息缓存
        fields.stream().filter(field -> field.isAnnotationPresent(Column.class)).forEach(this::parseColumn);
    }

    /**
     * 解析列
     *
     * @param field 对应的字段
     */
    private void parseColumn(Field field) {
        Column column = field.getAnnotation(Column.class);
        Integer index = fieldIndexMap.get(field);
        if (null != index && index > -1) {
            // 将column添加到map中缓存
            ExcelColumn ec = ExcelColumn.builder()
                .name(columnMap.get(index))
                .nullable(column.nullable())
                .wrap(column.wrap())
                .field(field)
                .build();

            ExcelUtils.checkColumn(ec, field);
            excelColumnMap.put(index, ec);
        }
    }

    /**
     * 处理换行
     *
     * @param text 原始文本
     * @return 处理后的文本
     */
    private String resolveWrap(String text) {
        return text.trim().replaceAll("\r?\n", "");
    }

    /**
     * 检验读取参数配置
     */
    private void validate() {
        if (null == workbook) {
            throw new IllegalStateException("Can not found workbook from this excel document.");
        }

        if (getSheetCount() <= 0) {
            throw new IllegalStateException("The excel document does not contains any sheet.");
        }

        if (config.getSheetIndex() >= getSheetCount()) {
            throw new IllegalStateException("The sheetIndex of ReaderConfig can not out of range 0 to " + (getSheetCount() - 1) + " that sheets count.");
        }
    }

    /**
     * 获取工作簿数量
     *
     * @return 工作簿数量
     */
    private int getSheetCount() {
        return workbook.getNumberOfSheets();
    }

    /**
     * 通过输入流创建工作簿
     *
     * @return 工作簿
     */
    private Workbook getWorkbook() {
        try {
            return WorkbookFactory.create(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 类型转换
     *
     * @param value  原始值
     * @param column 列信息
     * @return 转换后的值
     * @throws Exception 可能出现的异常
     */
    @SuppressWarnings("unchecked")
    private Object caseType(Object value, ExcelColumn column) throws Exception {
        if (null == value) {
            return null;
        }
        Class<?> type = column.getField().getType();

        if (type == String.class || type == CharSequence.class) {
            return String.valueOf(value);
        } else if (type == Integer.class || type == int.class) {
            return Integer.valueOf(emptyToZero(String.valueOf(value)));
        } else if (type == Float.class || type == float.class) {
            return Float.valueOf(emptyToZero(String.valueOf(value)));
        } else if (type == Byte.class || type == byte.class) {
            return Byte.valueOf(emptyToZero(String.valueOf(value)));
        } else if (type == Boolean.class || type == boolean.class) {
            return Boolean.valueOf(emptyToZero(String.valueOf(value)));
        } else if (type == Long.class || type == long.class) {
            return Long.valueOf(emptyToZero(String.valueOf(value)));
        } else if (type == Short.class || type == short.class) {
            return Short.valueOf(emptyToZero(String.valueOf(value)));
        } else if (type == Double.class || type == double.class) {
            return Double.valueOf(emptyToZero(String.valueOf(value)));
        } else if (type == Character.class || type == char.class) {
            String temp = String.valueOf(value);
            return temp.isEmpty() ? "" : temp.charAt(0);
        } else if (type == Date.class) {
            if (null != column.getParser()) {
                value = column.getParser().parse(value);
            } else {
                value = ExcelUtils.offeredDateParser().parse(value);
            }
        } else if (type == LocalDateTime.class) {
            if (null != column.getParser()) {
                value = column.getParser().parse(value);
            } else {
                value = ExcelUtils.offeredLocalDateTimeParser().parse(value);
            }
        } else if (type == java.sql.Date.class) {
            if (null != column.getParser()) {
                value = column.getParser().parse(value);
            } else {
                value = ExcelUtils.offeredSqlDateParser().parse(value);
            }
        } else if (type == Timestamp.class) {
            if (null != column.getParser()) {
                value = column.getParser().parse(value);
            } else {
                value = ExcelUtils.offeredTimestampParser().parse(value);
            }
        } else if (type == LocalDate.class) {
            // 自己处理吧
            if (null != column.getParser()) {
                value = column.getParser().parse(value);
            }
        } else if (type == LocalTime.class) {
            // 自己处理吧
            if (null != column.getParser()) {
                value = column.getParser().parse(value);
            }
        }
        return value;
    }

    /**
     * 将空字符串转换为 0
     *
     * @param text 字符串
     * @return 处理后的结果
     */
    private String emptyToZero(String text) {
        if (StringUtils.isEmpty(text)) {
            return "0";
        }
        return text;
    }
}
