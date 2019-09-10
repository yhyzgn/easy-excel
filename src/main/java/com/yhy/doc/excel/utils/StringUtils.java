package com.yhy.doc.excel.utils;

import java.util.regex.Pattern;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 15:47
 * version: 1.0.0
 * desc   : 字符串工具类
 */
public class StringUtils {

    private StringUtils() {
        throw new UnsupportedOperationException("Utils class can not be instantiate.");
    }

    public static boolean isEmpty(String text) {
        return null == text || text.trim().isEmpty();
    }

    public static boolean isNotEmpty(String text) {
        return !isEmpty(text);
    }

    public static String toUnicode(String text) {
        char[] chars = text.toCharArray();
        StringBuilder sb = new StringBuilder();
        for (char ch : chars) {
            sb.append("\\u").append(Integer.toString(ch, 16));
        }
        return sb.toString();
    }

    public static boolean isNumber(String text) {
        Pattern pattern = Pattern.compile("^[\\d]+(\\.[\\d]+)?$");
        return pattern.matcher(text).matches();
    }
}
