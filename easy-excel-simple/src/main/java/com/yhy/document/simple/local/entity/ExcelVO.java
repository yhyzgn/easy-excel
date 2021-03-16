package com.yhy.document.simple.local.entity;

import com.yhy.doc.excel.annotation.Converter;
import com.yhy.doc.excel.annotation.Excel;
import com.yhy.doc.excel.internal.EConverter;
import com.yhy.doc.excel.utils.StringUtils;
import lombok.Data;
import lombok.ToString;

import java.util.HashMap;
import java.util.Map;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2020-05-22 10:26 下午
 * version: 1.0.0
 * desc   :
 */
@Data
@ToString
public class ExcelVO {

    @Excel("区划编码")
    private String areaCode;

    @Excel(like = "业态%", export = "业态")
    @Converter(IndustryConverter.class)
    private Industry industry;

    @Excel(like = "企业名称%", export = "企业名称")
    private String name;

    @Excel(like = "统一社会信用代码%", export = "统一社会信用代码")
    private String creditCode;

    @Excel(like = "企业角色%", export = "企业角色")
    @Converter(StoreRoleConverter.class)
    private StoreRole storeRole;

    @Excel(like = "星级%", export = "星级")
    @Converter(GradeConverter.class)
    private Integer grade;

    @Excel(like = "营业住所地址%", export = "营业住所地址")
    private String address;

    @Excel(like = "企业申报人姓名%", export = "企业申报人姓名")
    private String reporterName;

    @Excel(like = "企业申报人手机号%", export = "企业申报人手机号")
    private String reporterMobile;

    public static class IndustryConverter implements EConverter<String, Industry> {
        @Override
        public Industry read(String s) {
            if (!StringUtils.isEmpty(s)) {
                // 规范的姿势匹配
                Industry industry = Industry.parse(s);
                if (null != industry) {
                    return industry;
                }

                // 万一没匹配到呢，万一又不规范呢 ...
                if (s.contains("景区")) {
                    return Industry.ScenicSpot;
                }
                if (s.contains("旅行社")) {
                    return Industry.TravelAgency;
                }
                if (s.contains("饭店") || s.contains("酒店")) {
                    return Industry.Hotel;
                }
                if (s.contains("歌舞") || s.contains("KTV")) {
                    return Industry.KTV;
                }
                if (s.contains("游戏") || s.contains("游艺")) {
                    return Industry.Game;
                }
                if (s.contains("演出")) {
                    return Industry.Show;
                }
                if (s.contains("互联网") || s.contains("上网") || s.contains("网吧")) {
                    return Industry.InternetBar;
                }
            }
            return null;
        }

        @Override
        public String write(Industry industry) {
            return industry.getName();
        }
    }

    public static class StoreRoleConverter implements EConverter<String, StoreRole> {
        @Override
        public StoreRole read(String s) {
            return StoreRole.parse(s);
        }

        @Override
        public String write(StoreRole storeRole) {
            return storeRole.getValue();
        }
    }

    public static class GradeConverter implements EConverter<String, Integer> {
        private static final Map<String, Integer> starTable = new HashMap<String, Integer>() {
            private static final long serialVersionUID = -5330751730938975576L;

            {
                put("一", 1);
                put("二", 2);
                put("三", 3);
                put("四", 4);
                put("五", 5);
            }
        };

        @Override
        public Integer read(String s) {
            if (null != s && !s.isEmpty()) {
                s = s.trim().toUpperCase();
                if ("无".equals(s)) {
                    return null;
                }
                // 星级解析
                if (s.matches("^A+.*")) {
                    // A AA AAA AAAA AAAAA ... 等级
                    return s.length();
                }
                if (s.matches("^\\dA.*")) {
                    // 1A 2A 3A 4A 5A ... 等级
                    return charToInteger(s.charAt(0));
                }
                if (s.matches("^\\d星.*")) {
                    // 1星 2星 3星 4星 5星 ... 等级
                    return charToInteger(s.charAt(0));
                }
                if (s.matches("^[一二三四五]星.*")) {
                    // 一星 二星 三星 四星 五星 ... 等级
                    return starTable.get(String.valueOf(s.charAt(0)));
                }
            }
            return null;
        }

        @Override
        public String write(Integer integer) {
            return null == integer ? "" : integer.toString();
        }

        private Integer charToInteger(char ch) {
            return Integer.parseInt(String.valueOf(ch));
        }
    }
}