package com.yhy.doc.excel.internal;

import lombok.Data;
import lombok.ToString;
import lombok.experimental.Accessors;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:51
 * version: 1.0.0
 * desc   :
 */
@Data
@ToString
@Accessors(chain = true)
public class ReaderConfig {

    private int sheetIndex;

    private int titleIndex;

    private int rowStartIndex;

    private int cellStartIndex;
}
