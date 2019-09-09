package com.yhy.doc.excel.io;

import com.yhy.doc.excel.internal.ReaderConfig;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
    private Class<T> clazz;
    private int sheetCount;
    private int sheetIndex;

    private ExcelReader(InputStream is, ReaderConfig config) {
        this.is = is;
        this.config = config;
        this.workbook = getWorkbook();
        this.titleMap = new HashMap<>();
        this.valueList = new ArrayList<>();
        validate();
    }

    @SuppressWarnings("unchecked")
    public static <T> ExcelReader create(InputStream is, ReaderConfig config) {
        return new ExcelReader(is, config);
    }

    public void read(@NonNull Class<T> clazz) {
        this.clazz = clazz;
        this.sheetCount = getSheetCount();
        this.sheetIndex = config.getSheetIndex();
        // 开始读取
        reading();
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
                    // 日期
                    value = cell.getDateCellValue().toString();
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
    }

    private void parseTitles() {
        if (null == clazz) {
            throw new IllegalArgumentException("Must set the class of reading result with method read(Class<?> clazz).");
        }

        Field[] fields = clazz.getDeclaredFields();
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
}
