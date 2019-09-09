package com.yhy.doc.excel;

import com.yhy.doc.excel.internal.ReaderConfig;
import com.yhy.doc.excel.io.ExcelReader;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:42
 * version: 1.0.0
 * desc   :
 */
@Slf4j
public class Tester {

    @Test
    public void test() throws FileNotFoundException {
//        double similarity = CosineSimilarity.getSimilarity("统一信用代码/工商注册号", "统一信用代码");
//        log.info("相似度：{}", similarity);

        ReaderConfig config = new ReaderConfig();
        config.setSheetIndex(0).setTitleIndex(1).setCellStartIndex(1);
        ExcelReader reader = ExcelReader.create(new FileInputStream("F:\\aa.xlsx"), config);
        reader.read(null);
    }
}
