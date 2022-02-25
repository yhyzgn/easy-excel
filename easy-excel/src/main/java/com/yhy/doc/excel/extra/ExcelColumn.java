package com.yhy.doc.excel.extra;

import com.yhy.doc.excel.internal.EConverter;
import com.yhy.doc.excel.internal.EDateParser;
import com.yhy.doc.excel.internal.EFilter;
import lombok.*;

import java.lang.reflect.Field;

/**
 * 表列信息
 * <p>
 * Created on 2019-09-09 21:19
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@SuppressWarnings("rawtypes")
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelColumn {

    /**
     * 表头名称
     */
    @NonNull
    private String name;

    /**
     * 是否可为空
     */
    private boolean nullable;

    /**
     * 是否自动处理换行符
     */
    private boolean wrap;

    /**
     * 标题对应的字段
     */
    private Field field;

    /**
     * 计算公式
     */
    private String formula;

    /**
     * 合并单元格范围
     */
    private Rect mergeRect;

    /**
     * 过滤器
     */
    private EFilter filter;

    /**
     * 转换器
     */
    private EConverter converter;

    /**
     * 日期解析器
     */
    private EDateParser parser;

    /**
     * 行高
     */
    private short rowHeight;
}
