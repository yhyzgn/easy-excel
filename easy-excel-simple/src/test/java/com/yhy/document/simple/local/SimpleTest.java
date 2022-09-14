package com.yhy.document.simple.local;

import com.yhy.doc.excel.Excel;
import com.yhy.doc.excel.extra.ReaderConfig;
import com.yhy.document.simple.local.entity.Company;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.util.List;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:42
 * version: 1.0.0
 * desc   :
 */
@Slf4j
@SpringBootTest
public class SimpleTest {

    @Test
    public void test() throws Exception {
//        System.out.println(StringUtils.isEmail("1436433211@qq.com"));

//        String text = "abc\r\ndef\nfdfwer";
//        System.out.println(text.replaceAll("\r?\n", ""));

//        System.out.println(StringUtils.isNumber("122.23"));

//        List<Integer> test = new ArrayList<>();
//        test.add(1);
//        test.add(2);
//        test.add(3);
//        test.add(6);
//        test.add(6);
//        test.add(3);
//        test.add(4);
//
//        test.stream().sorted(Comparator.comparingInt(o -> o)).forEach(System.out::println);

        ReaderConfig config = new ReaderConfig()
            .setSheetIndex(0)
            .setTitleIndex(0);

        ClassPathResource res = new ClassPathResource("easy-excel-simple.xlsx");
        List<Company> companyList = Excel.read(res.getInputStream(), config, Company.class);
        companyList.forEach(company -> {
            log.info(company.toString());
        });

//        List<ExcelVO> voList = ExcelUtils.read(new File("/Users/yhyzgn/Downloads/vv.xlsx"), config, ExcelVO.class);
//        voList = ExcelUtils.filter(voList, vo -> null != vo.getCreditCode() && !vo.getCreditCode().isEmpty());
//        voList.forEach(vo -> {
//            log.info(vo.toString());
//        });

        Excel.write(new File("/Users/yhyzgn/Downloads/easy-excel-simple-out.xlsx"), companyList, "测试看看", Company.class);
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
}
