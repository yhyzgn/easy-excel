package com.yhy.doc.excel.io;

import com.yhy.doc.excel.annotation.Converter;
import com.yhy.doc.excel.annotation.Excel;
import com.yhy.doc.excel.annotation.Filter;
import com.yhy.doc.excel.annotation.Formatter;
import com.yhy.doc.excel.ers.ExcelConverter;
import com.yhy.doc.excel.ers.ExcelFilter;
import com.yhy.doc.excel.ers.ExcelFormatter;
import com.yhy.doc.excel.internal.CosineSimilarity;
import com.yhy.doc.excel.internal.ExcelTitle;
import com.yhy.doc.excel.internal.ReaderConfig;
import com.yhy.doc.excel.utils.ExcelUtils;
import com.yhy.doc.excel.utils.StringUtils;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.regex.Pattern;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:41
 * version: 1.0.0
 * desc   :
 */
@Slf4j
public class ExcelReader<T> {
    private InputStream is;
    private ReaderConfig config;
    private Workbook workbook;
    private Map<Integer, String> titleMap;
    private List<Map<Integer, Object>> valueList;
    private Map<Integer, ExcelTitle> excelTitleMap;
    private Class<T> clazz;
    private Constructor<T> constructor;
    private int sheetCount;
    private int sheetIndex;
    private List<T> resultList;

    private ExcelReader(InputStream is, ReaderConfig config) {
        this.is = is;
        this.config = config;
        this.workbook = getWorkbook();
        this.titleMap = new HashMap<>();
        this.valueList = new ArrayList<>();
        this.excelTitleMap = new HashMap<>();
        validate();
    }

    @SuppressWarnings("unchecked")
    public static <T> ExcelReader<T> create(InputStream is, ReaderConfig config) {
        return new ExcelReader(is, config);
    }

    public List<T> read(@NonNull Class<T> clazz) {
        if (null == clazz) {
            throw new IllegalArgumentException("The argument clazz can not be null.");
        }
        try {
            constructor = clazz.getConstructor();
        } catch (NoSuchMethodException e) {
            throw new IllegalArgumentException("Your model class '" + clazz.getName() + "' must contains a constructor without any argument, but not found.");
        }
        this.clazz = clazz;
        this.sheetCount = getSheetCount();
        this.sheetIndex = config.getSheetIndex();
        // 开始读取
        reading();
        return resultList;
    }

    private void reading() {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int rows = sheet.getPhysicalNumberOfRows();
        if (config.getTitleIndex() >= rows) {
            throw new IllegalStateException("The titleIndex of ReaderConfig is error.");
        }

        Row title = sheet.getRow(config.getTitleIndex());
        readRow(title, true);

        // 开始行的索引，不设置的话，默认从标题的下一行开始
        int start = config.getRowStartIndex() > 0 ? config.getRowStartIndex() : config.getTitleIndex() + 1;
        if (start > rows) {
            throw new IllegalStateException("The rowStartIndex of ReaderConfig is error.");
        }

        Row row;
        for (int i = start; i < rows; i++) {
            row = sheet.getRow(i);
            readRow(row, false);
        }
        parse();
    }

    private void readRow(Row row, boolean isTitle) {
        // 列的开始索引
        int start = config.getCellStartIndex();
        if (null != row) {
            Cell cell;
            String value;
            Map<Integer, Object> valuesOfRow = null;
            int cells = row.getPhysicalNumberOfCells();
            if (!isTitle) {
                // 不是标题
                valuesOfRow = new HashMap<>();
            }
            for (int i = start; i < cells; i++) {
                cell = row.getCell(i);
                if (null != cell) {
                    value = getValueOfCell(cell);
                    // 标题，添加到标题map中
                    if (isTitle) {
                        // 第j列的标题
                        titleMap.put(i, value);
                    } else {
                        // 往下其他行，表格值
                        valuesOfRow.put(i, value);
                    }
                }
            }
            if (null != valuesOfRow) {
                valueList.add(valuesOfRow);
            }
        }
    }

    private String getValueOfCell(Cell cell) {
        //判断是否为null或空串
        if (null == cell || "".equals(cell.toString().trim())) {
            return "";
        }
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        String value;
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            type = evaluator.evaluate(cell).getCellType();
        }
        switch (type) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    // 日期时间，转换为毫秒
                    value = String.valueOf(cell.getDateCellValue().getTime());
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
        return value;
    }

    private void parse() {
        if (titleMap.isEmpty()) {
            throw new IllegalStateException("Can not read titles of excel file.");
        }

        parseTitles();

        parseData();
    }

    private void parseData() {
        if (!valueList.isEmpty()) {
            resultList = new ArrayList<>();
            valueList.stream().filter(item -> {
                ExcelTitle title;
                for (Map.Entry<Integer, Object> et : item.entrySet()) {
                    title = excelTitleMap.get(et.getKey());
                    // 如果不允许为空缺读取到空值，则忽略该行
                    if (null == title || null == et.getValue() && !title.isNullable()) {
                        return false;
                    }
                }
                return true;
            }).forEach(item -> {
                try {
                    T data = constructor.newInstance();
                    Integer index;
                    Object value;
                    ExcelTitle title;
                    Method setter;
                    for (Map.Entry<Integer, Object> et : item.entrySet()) {
                        index = et.getKey();
                        value = et.getValue();
                        value = String.valueOf(value);
                        title = excelTitleMap.get(index);

                        // 自动处理换行符
                        if (title.isWrap()) {
                            value = resolveWrap(String.valueOf(value));
                        }

                        // 先执行过滤器
                        if (null != title.getFilter()) {
                            value = title.getFilter().read(String.valueOf(value));
                        }

                        // 执行转换器，格式化一些值得转换，比如枚举
                        // 两者同时设置的话，前者生效
                        if (null != title.getConverter()) {
                            value = title.getConverter().read(String.valueOf(value));
                        } else if (null != title.getFormatter()) {
                            value = title.getFormatter().read(String.valueOf(value));
                        }

                        // 类型转换
                        value = caseType(value, title.getField().getType());

//                        if (!title.getField().getType().isAssignableFrom(value.getClass())) {
//                            // 转换后类型还是不匹配
//                            throw new IllegalStateException("The type '" + value.getClass() + "' of value '" + value.toString() + "' can not set to field '" + title.getField().getName() + "' what type is '" + title.getField().getType() + "'.");
//                        }

                        // 字段对应的setter方法
                        setter = ExcelUtils.setter(title.getField(), clazz);
                        // 执行getter方法，设置值
                        setter.invoke(data, value);
                    }
                    resultList.add(data);
                } catch (Exception e) {
                }
            });
        }
    }

    private void parseTitles() {
        List<Field> fields = new ArrayList<>(Arrays.asList(clazz.getDeclaredFields()));
        // 将标题信息缓存
        fields.stream().filter(field -> field.isAnnotationPresent(Excel.class)).forEach(this::parseTitle);
    }

    private void parseTitle(Field field) {
        Excel excel = field.getAnnotation(Excel.class);
        // 先进行name完全匹配
        int index = indexOfTitle(excel.value(), excel.insensitive());
        // 未匹配到正确的标题，进行模糊匹配
        if (index == -1) {
            index = indexOfTitleByLike(excel.like(), excel.insensitive());
        }
        // 如果还是未匹配到并且开启了智能匹配，就进行智能匹配
        if (index == -1 && excel.intelligent()) {
            index = indexOfTitleByIntelligent(excel.value(), excel.insensitive(), excel.tolerance());
        }

        // 如果真还是没找到，那就是天命了，只能忽略了...
        if (index > -1) {
            // 将title添加到map中缓存
            ExcelTitle title = new ExcelTitle(titleMap.get(index)).setNullable(excel.nullable()).setWrap(excel.wrap()).setField(field);
            // 扫描过滤器
            Filter filter = field.getAnnotation(Filter.class);
            if (null != filter && filter.value() != ExcelFilter.class) {
                title.setFilter(ExcelUtils.instantiate(filter.value()));
            }
            // 扫描转换器
            Converter converter = field.getAnnotation(Converter.class);
            if (null != converter && converter.value() != ExcelConverter.class) {
                title.setConverter(ExcelUtils.instantiate(converter.value()));
            }
            // 扫描格式化模式
            Formatter formatter = field.getAnnotation(Formatter.class);
            if (null != formatter && formatter.value() != ExcelFormatter.class) {
                title.setFormatter(ExcelUtils.instantiate(formatter.value()));
            }
            excelTitleMap.put(index, title);
        }
    }

    private int indexOfTitleByIntelligent(String name, boolean insensitive, double tolerance) {
        name = name.trim();
        if (insensitive) {
            name = name.toLowerCase(Locale.getDefault());
        }
        String title;
        for (Map.Entry<Integer, String> et : titleMap.entrySet()) {
            title = resolveWrap(et.getValue());
            if (insensitive) {
                title = title.toLowerCase(Locale.getDefault());
            }

            // 求得相似度
            double similarity = CosineSimilarity.getSimilarity(name, title);
            if (similarity >= 1.0 - tolerance) {
                // 相似度在容差范围以内，表示匹配成功
                return et.getKey();
            }
        }
        return -1;
    }

    private int indexOfTitleByLike(String like, boolean insensitive) {
        // 正则表达式，将 % 转换为 .*?
        like = like.trim().replaceAll("%+", ".*?");
        Pattern pattern = insensitive ? Pattern.compile(like, Pattern.CASE_INSENSITIVE) : Pattern.compile(like);
        String title;
        for (Map.Entry<Integer, String> et : titleMap.entrySet()) {
            title = resolveWrap(et.getValue());
            if (pattern.matcher(title).matches()) {
                return et.getKey();
            }
        }
        return -1;
    }

    private int indexOfTitle(String name, boolean insensitive) {
        name = name.trim();
        String title;
        for (Map.Entry<Integer, String> et : titleMap.entrySet()) {
            title = resolveWrap(et.getValue());
            if (insensitive) {
                // 忽略大小写
                if (name.equalsIgnoreCase(title)) {
                    return et.getKey();
                }
            } else {
                // 严格大小写
                if (name.equals(title)) {
                    return et.getKey();
                }
            }
        }
        return -1;
    }

    private String resolveWrap(String text) {
        return text.trim().replace("\n", "");
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

    private Object caseType(Object value, Class<?> type) {
        // 如果是日期类型
        if (type == Date.class || type == java.sql.Date.class || type == Timestamp.class || type == LocalDate.class || type == LocalDateTime.class) {
            return value;
        }
        if (StringUtils.isEmpty(String.valueOf(value))) {
            return value;
        }
        if (type == String.class) {
            return String.valueOf(value);
        } else if (type == Integer.class || type == int.class) {
            return Integer.valueOf(String.valueOf(value));
        } else if (type == Float.class || type == float.class) {
            return Float.valueOf(String.valueOf(value));
        } else if (type == Byte.class || type == byte.class) {
            return Byte.valueOf(String.valueOf(value));
        } else if (type == Boolean.class || type == boolean.class) {
            return Boolean.valueOf(String.valueOf(value));
        } else if (type == Long.class || type == long.class) {
            return Long.valueOf(String.valueOf(value));
        } else if (type == Short.class || type == short.class) {
            return Short.valueOf(String.valueOf(value));
        } else if (type == Double.class || type == double.class) {
            return Double.valueOf(String.valueOf(value));
        } else if (type == Character.class || type == char.class) {
            return String.valueOf(value).charAt(0);
        }
        return value;
    }
}
