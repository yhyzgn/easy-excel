package com.yhy.doc.excel.extra;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.ToString;
import lombok.experimental.Accessors;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 11:58
 * version: 1.0.0
 * desc   : 单元格位置
 */
@Data
@ToString
@AllArgsConstructor
@Accessors(chain = true)
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
