package com.yhy.doc.excel;

import com.yhy.doc.excel.entity.Company;
import com.yhy.doc.excel.ers.ExcelConverter;
import com.yhy.doc.excel.internal.ReaderConfig;
import com.yhy.doc.excel.io.ExcelReader;
import com.yhy.doc.excel.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.List;

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
    public void test() throws Exception {
//        System.out.println(StringUtils.isNumber("122.23"));

        ReaderConfig config = new ReaderConfig();
        config.setSheetIndex(0).setTitleIndex(1).setCellStartIndex(0);
        ExcelReader<Company> reader = ExcelReader.create(new FileInputStream("F:\\aa.xlsx"), config);
        List<Company> companyList = reader.read(Company.class);
        companyList.forEach(company -> {
            log.info(company.toString());
        });


//        Type type = ExcelUtils.getParamType(TestInterface.class, ExcelConverter.class, 0);
//        System.out.println(Class.forName(type.getTypeName()));

//        System.out.println(CharSequence.class.isAssignableFrom(String.class));

//        String like = "%你们%好%%啊".replaceAll("%+", ".*?");
//        log.info(like);

//        double similarity = CosineSimilarity.getSimilarity("统一信用代码/工商注册号", "统一信用代码");
//        log.info("相似度：{}", similarity);

//        String unicode = StringUtils.toUnicode("相似度");
//        log.info(unicode);

//        boolean test = Pattern.compile(".*度.*?你.*?哈哈").matcher("相似度aert你们的沪电股份的哈哈").matches();
//        log.info(test + "");

//        ReaderConfig config = new ReaderConfig();
//        config.setSheetIndex(0).setTitleIndex(1).setCellStartIndex(1);
//        ExcelReader reader = ExcelReader.create(new FileInputStream("F:\\aa.xlsx"), config);
//        reader.read(null);
    }

    public static class TestInterface implements ExcelConverter<Integer> {

        @Override
        public Integer read(Object value) {
            return null;
        }

        @Override
        public Object write(Integer value) {
            return null;
        }
    }
}
