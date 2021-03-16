package com.yhy.document.simple.local.entity;

import lombok.Getter;
import lombok.ToString;

import java.util.Arrays;
import java.util.Optional;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 16:06
 * version: 1.0.0
 * desc   :
 */
@Getter
@ToString
public enum Sex {
    Male(0, "男"), Female(1, "女"), Secret(2, "保密");

    private Integer code;

    private String value;

    Sex(Integer code, String value) {
        this.code = code;
        this.value = value;
    }

    public static Sex parse(Integer code) {
        Optional<Sex> optional = Arrays.stream(Sex.values()).filter(sex -> sex.code.equals(code)).findFirst();
        return optional.orElse(null);
    }

    public static Sex parse(String value) {
        Optional<Sex> optional = Arrays.stream(Sex.values()).filter(sex -> sex.value.equals(value)).findFirst();
        return optional.orElse(null);
    }
}
