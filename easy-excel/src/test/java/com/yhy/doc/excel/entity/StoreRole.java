package com.yhy.doc.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.ToString;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-22 10:27 下午
 * version: 1.0.0
 * desc   :
 */
@Getter
@ToString
@AllArgsConstructor
public enum StoreRole {
    /**
     * 总店
     */
    Master(1, "总店"),

    /**
     * 分店
     */
    Branch(2, "分店"),

    /**
     * 网点
     */
    Internet(3, "网点"),
    ;

    private final Integer code;
    private final String value;

    public static StoreRole parse(Integer code) {
        for (StoreRole sr : StoreRole.values()) {
            if (sr.code.equals(code)) {
                return sr;
            }
        }
        return null;
    }

    public static StoreRole parse(String value) {
        for (StoreRole sr : StoreRole.values()) {
            if (sr.value.equals(value)) {
                return sr;
            }
        }
        return null;
    }
}
