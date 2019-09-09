package com.yhy.doc.excel.internal;

import com.yhy.doc.excel.ers.ExcelConverter;
import com.yhy.doc.excel.ers.ExcelFilter;
import com.yhy.doc.excel.ers.ExcelFormatter;
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
public class ExcelTitle {

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
     * 过滤器
     */
    private ExcelFilter filter;

    /**
     * 转换器
     */
    private ExcelConverter converter;

    /**
     * 格式化
     */
    private ExcelFormatter formatter;
}
