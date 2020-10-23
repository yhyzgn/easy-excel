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
public enum Industry {
    /**
     * A级景区
     */
    ScenicSpot(1, "A级景区"),

    /**
     * 旅行社
     */
    TravelAgency(2, "旅行社"),

    /**
     * 星级饭店
     */
    Hotel(3, "星级饭店"),

    /**
     * 歌舞娱乐场所（KTV）
     */
    KTV(4, "歌舞娱乐场所"),

    /**
     * 游戏游艺场所
     */
    Game(5, "游戏游艺场所"),

    /**
     * 营业性演出活动
     */
    Show(6, "营业性演出活动"),

    /**
     * 互联网上网服务营业场所（网吧）
     */
    InternetBar(7, "互联网上网服务营业场所（网吧）"),
    ;

    private final Integer code;
    private final String name;

    public static Industry parse(Integer code) {
        for (Industry idt : Industry.values()) {
            if (idt.code.equals(code)) {
                return idt;
            }
        }
        return null;
    }

    public static Industry parse(String name) {
        for (Industry idt : Industry.values()) {
            if (idt.name.equals(name)) {
                return idt;
            }
        }
        return null;
    }
}
