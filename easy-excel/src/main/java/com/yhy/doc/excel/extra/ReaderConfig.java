package com.yhy.doc.excel.extra;

import lombok.*;
import lombok.experimental.Accessors;

/**
 * 读取文件配置类
 * <p>
 * Created on 2019-09-09 12:51
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
@Data
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

    /**
     * 默认配置
     *
     * @return 默认
     */
    public static ReaderConfig deft() {
        return new ReaderConfig();
    }
}
