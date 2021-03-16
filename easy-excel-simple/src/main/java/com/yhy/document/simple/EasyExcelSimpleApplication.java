package com.yhy.document.simple;

import com.yhy.document.simple.local.Tester;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class EasyExcelSimpleApplication {

    public static void main(String[] args) {
        try {
            new Tester().test();
        } catch (Exception e) {
            e.printStackTrace();
        }

//        SpringApplication.run(EasyExcelSimpleApplication.class, args);
    }
}
