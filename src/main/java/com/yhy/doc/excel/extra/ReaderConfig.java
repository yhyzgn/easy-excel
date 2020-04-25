package com.yhy.doc.excel.extra;

import lombok.Data;
import lombok.ToString;
import lombok.experimental.Accessors;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:51
 * version: 1.0.0
 * desc   : 读取文件配置类
 */
@Data
@ToString
@Accessors(chain = true)
public class ReaderConfig {

    /**
     * sheet文档索引，从0开始计
     */
    private int sheetIndex = 0;

    /**
     * 标题栏行索引，从0开始计
     */
    private int titleIndex = 0;

    /**
     * 内容从第几行开始读取，从0开始计
     */
    private int rowStartIndex = 0;

    /**
     * 内容到第几行读取结束，从0开始计
     */
    private int rowEndIndex = -1;

    /**
     * 内容从第几列开始读取，从0开始计
     */
    private int cellStartIndex = 0;

    /**
     * 内容到第几列读取结束，从0开始计
     */
    private int cellEndIndex = -1;
}
