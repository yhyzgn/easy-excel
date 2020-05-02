package com.yhy.doc.excel.extra;

import com.yhy.doc.excel.annotation.Style;
import com.yhy.doc.excel.internal.EConverter;
import com.yhy.doc.excel.internal.EDateParser;
import com.yhy.doc.excel.internal.EFilter;
import lombok.*;
import lombok.experimental.Accessors;

import java.lang.reflect.Field;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 21:19
 * version: 1.0.0
 * desc   : 表头头部信息
 */
@Data
@ToString
@NoArgsConstructor
@RequiredArgsConstructor
@Accessors(chain = true)
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
