package com.yhy.doc.excel.utils;

import org.apache.poi.common.usermodel.HyperlinkType;

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

    public static boolean isEmail(String text) {
        Pattern pattern = Pattern.compile("^(mailto:)?[A-Za-z0-9\\u4e00-\\u9fa5]+@[a-zA-Z0-9_-]+(\\.[a-zA-Z0-9_-]+)+$");
        return pattern.matcher(text).matches();
    }

    public static boolean isFile(String text) {
        // false for moment.
        return false;
    }

    private static boolean isUrl(String text) {
        return text.startsWith("https://") || text.startsWith("http://") || text.startsWith("ws://") || text.startsWith("file://") || text.startsWith("ftp://");
    }

    public static boolean isHyperLink(String text) {
        return isEmail(text) || isUrl(text) || isFile(text);
    }

    public static HyperlinkType hyperLinkType(String text) {
        if (isEmail(text)) {
            return HyperlinkType.EMAIL;
        } else if (isUrl(text)) {
            return HyperlinkType.URL;
        } else if (isFile(text)) {
            return HyperlinkType.FILE;
        } else {
            return HyperlinkType.NONE;
        }
    }
}
