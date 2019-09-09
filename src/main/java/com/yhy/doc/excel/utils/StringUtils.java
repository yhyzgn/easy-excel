package com.yhy.doc.excel.utils;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 15:47
 * version: 1.0.0
 * desc   :
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
}
