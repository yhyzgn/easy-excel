package com.yhy.doc.excel.extra;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 单元格位置
 * <p>
 * Created on 2019-09-10 11:58
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Rect {

    /**
     * 是否被合并
     */
    private boolean merged;

    /**
     * 单元格开始行索引
     */
    private int rowStart;

    /**
     * 单元格结束行索引
     */
    private int rowEnd;

    /**
     * 单元格开始列索引
     */
    private int columnStart;

    /**
     * 单元格结束列索引
     */
    private int columnEnd;
}
