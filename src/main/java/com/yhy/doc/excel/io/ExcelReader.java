package com.yhy.doc.excel.io;

import com.yhy.doc.excel.annotation.Excel;
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
import java.lang.reflect.Method;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;
import java.util.regex.Pattern;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:41
 * version: 1.0.0
 * desc   : Excel读取器
 */
@Slf4j
public class ExcelReader<T> {
    private final InputStream is;
    private final ReaderConfig config;
    private final Workbook workbook;
    private Map<Integer, String> columnMap;
    private List<Map<Integer, Object>> valueList;
    private Map<Integer, ExcelColumn> excelColumnMap;
    private Class<T> clazz;
    private Sheet sheet;
    private Constructor<T> constructor;
    private int sheetIndex;
    private List<T> resultList;

    public ExcelReader(File file, ReaderConfig config) throws FileNotFoundException {
        this(new FileInputStream(file), config);
    }

    public ExcelReader(ServletRequest request, ReaderConfig config) throws IOException {
        this(request.getInputStream(), config);
    }

    public ExcelReader(InputStream is, ReaderConfig config) {
        this.is = is;
        this.config = config;
        this.workbook = getWorkbook();
        this.columnMap = new HashMap<>();
        this.valueList = new ArrayList<>();
        this.excelColumnMap = new HashMap<>();
        validate();
    }

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

    public void release() throws IOException {
        if (null != is) {
            is.close();
        }
        if (null != workbook) {
            workbook.close();
        }
        columnMap = null;
        valueList = null;
        excelColumnMap = null;
        sheet = null;
        resultList = null;
    }

    private void reading() {
        sheet = workbook.getSheetAt(sheetIndex);
        // sheet.getPhysicalNumberOfRows() 方法获取到的行数会自动忽略合并的单元格
        int lastRowIndex = config.getRowEndIndex() > -1 ? config.getRowEndIndex() : sheet.getLastRowNum();
        // 开始行的索引，不设置的话，默认从标题的下一行开始
        int firstRowIndex = config.getRowStartIndex() > 0 ? config.getRowStartIndex() : config.getTitleIndex() + 1;
        int rows = lastRowIndex - firstRowIndex + 1;

        // 读取标题
        readColumn();

        // 读取其他行
        readRows(firstRowIndex, rows);

        parse();
    }

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
                // row.getLastCellNum() 结果也是 合并单元格只算1格，所以合并单元格的值还要手动判断获取
                int cells = config.getCellEndIndex() > -1 ? config.getCellEndIndex() : row.getPhysicalNumberOfCells();
                for (int j = columnStart; j < cells + columnStart; j++) {
                    cell = row.getCell(j);
                    if (null != cell) {
                        value = getValueOfCell(cell, false);
                        // 往下其他行，表格值
                        rect = ExcelUtils.merged(sheet, i, j, rowStart, columnStart);
                        if (rect.isMerged() && rect.getColumnStart() < rect.getColumnEnd()) {
                            // 在合并的单元格内，大单元格内的所有小单元格都设置同一个值
                            // 列
                            if (j == rect.getColumnStart()) {
                                // 合并单元格内
                                for (int k = j; k <= rect.getColumnEnd(); k++) {
                                    valuesOfRow.put(k, value);
                                }
                            } else {
                                // 合并单元格之后的单元格索引，需要原来的i加上大单元格所占的最后索引
                                valuesOfRow.put(j + rect.getColumnEnd(), value);
                            }
                        } else {
                            valuesOfRow.put(j, value);
                        }
                    }
                }
                valueList.add(valuesOfRow);
            }
        }
    }

    private void readColumn() {
        Row column = sheet.getRow(config.getTitleIndex());
        // 列的开始索引
        if (null != column) {
            Cell cell;
            Rect rect;
            Object value;
            int start = config.getCellStartIndex();
            int cells = column.getPhysicalNumberOfCells();
            for (int j = start; j < cells; j++) {
                cell = column.getCell(j);
                if (null != cell) {
                    value = getValueOfCell(cell, true);
                    // 标题，添加到标题map中
                    // 第j列的标题
                    columnMap.put(j, String.valueOf(value));
                }
            }
        }
    }

    private Object getValueOfCell(Cell cell, boolean isTitle) {
        //判断是否为null或空串
        if (null == cell || "".equals(cell.toString().trim())) {
            return null;
        }
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Object value;
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            type = evaluator.evaluate(cell).getCellType();
        }
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

    private void parse() {
        if (columnMap.isEmpty()) {
            throw new IllegalStateException("Can not read columns of excel file.");
        }

        parseColumns();

        parseData();
    }

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
                    Method setter;
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

    private void parseColumns() {
        List<Field> fields = new ArrayList<>(Arrays.asList(clazz.getDeclaredFields()));
        // 将标题信息缓存
        fields.stream().filter(field -> field.isAnnotationPresent(Excel.class)).forEach(this::parseColumn);
    }

    private void parseColumn(Field field) {
        Excel excel = field.getAnnotation(Excel.class);
        // 先进行name完全匹配
        int index = indexOfColumn(excel.value(), excel.insensitive());
        // 未匹配到正确的标题，进行模糊匹配
        if (index == -1) {
            index = indexOfColumnByLike(excel.like(), excel.insensitive());
        }
        // 如果还是未匹配到并且开启了智能匹配，就进行智能匹配
        if (index == -1 && excel.intelligent()) {
            index = indexOfColumnByIntelligent(excel.value(), excel.insensitive(), excel.tolerance());
        }

        // 如果真还是没找到，那就是天命了，只能忽略了...
        if (index > -1) {
            // 将column添加到map中缓存
            ExcelColumn column = new ExcelColumn(columnMap.get(index)).setNullable(excel.nullable()).setWrap(excel.wrap()).setField(field);
            ExcelUtils.checkColumn(column, field);
            excelColumnMap.put(index, column);
        }
    }

    private int indexOfColumnByIntelligent(String name, boolean insensitive, double tolerance) {
        name = name.trim();
        if (insensitive) {
            name = name.toLowerCase(Locale.getDefault());
        }
        String column;
        for (Map.Entry<Integer, String> et : columnMap.entrySet()) {
            column = resolveWrap(et.getValue());
            if (insensitive) {
                column = column.toLowerCase(Locale.getDefault());
            }

            // 求得相似度
            double similarity = CosineSimilarity.getSimilarity(name, column);
            if (similarity >= 1.0D - tolerance) {
                // 相似度在容差范围以内，表示匹配成功
                return et.getKey();
            }
        }
        return -1;
    }

    private int indexOfColumnByLike(String like, boolean insensitive) {
        // 正则表达式，将 % 转换为 .*?
        like = like.trim().replaceAll("%+", ".*?");
        Pattern pattern = insensitive ? Pattern.compile(like, Pattern.CASE_INSENSITIVE) : Pattern.compile(like);
        String column;
        for (Map.Entry<Integer, String> et : columnMap.entrySet()) {
            column = resolveWrap(et.getValue());
            if (pattern.matcher(column).matches()) {
                return et.getKey();
            }
        }
        return -1;
    }

    private int indexOfColumn(String name, boolean insensitive) {
        name = name.trim();
        String column;
        for (Map.Entry<Integer, String> et : columnMap.entrySet()) {
            column = resolveWrap(et.getValue());
            if (insensitive) {
                // 忽略大小写
                if (name.equalsIgnoreCase(column)) {
                    return et.getKey();
                }
            } else {
                // 严格大小写
                if (name.equals(column)) {
                    return et.getKey();
                }
            }
        }
        return -1;
    }

    private String resolveWrap(String text) {
        return text.trim().replaceAll("\r?\n", "");
    }

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

    private int getSheetCount() {
        return workbook.getNumberOfSheets();
    }

    private Workbook getWorkbook() {
        try {
            return WorkbookFactory.create(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    @SuppressWarnings("unchecked")
    private Object caseType(Object value, ExcelColumn column) throws Exception {
        if (null == value) {
            return null;
        }
        Class<?> type = column.getField().getType();

        if (type == String.class || type == CharSequence.class) {
            return String.valueOf(value);
        } else if (type == Integer.class || type == int.class) {
            return Integer.valueOf(emptyOrZero(String.valueOf(value)));
        } else if (type == Float.class || type == float.class) {
            return Float.valueOf(emptyOrZero(String.valueOf(value)));
        } else if (type == Byte.class || type == byte.class) {
            return Byte.valueOf(emptyOrZero(String.valueOf(value)));
        } else if (type == Boolean.class || type == boolean.class) {
            return Boolean.valueOf(emptyOrZero(String.valueOf(value)));
        } else if (type == Long.class || type == long.class) {
            return Long.valueOf(emptyOrZero(String.valueOf(value)));
        } else if (type == Short.class || type == short.class) {
            return Short.valueOf(emptyOrZero(String.valueOf(value)));
        } else if (type == Double.class || type == double.class) {
            return Double.valueOf(emptyOrZero(String.valueOf(value)));
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

    private String emptyOrZero(String text) {
        if (!StringUtils.isNumber(text) || text.trim().isEmpty()) {
            return "0";
        }
        return text;
    }
}
