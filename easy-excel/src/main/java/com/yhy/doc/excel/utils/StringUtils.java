package com.yhy.doc.excel.utils;

import org.apache.poi.common.usermodel.HyperlinkType;

import java.util.regex.Pattern;

/**
 * Excel 字符串工具类
 * <p>
 * Created on 2019-09-09 15:47
 *
 * @author 颜洪毅
 * @version 1.0.0
 * @since 1.0.0
 */
public class StringUtils {

    /**
     * 禁止实例化处理
     */
    private StringUtils() {
        throw new UnsupportedOperationException("Utils class can not be instantiate.");
    }

    /**
     * 是否是空字符串
     *
     * @param text 字符串
     * @return 是否空
     */
    public static boolean isEmpty(String text) {
        return null == text || text.trim().isEmpty();
    }

    /**
     * 是否非空
     *
     * @param text 字符串
     * @return 是否非空
     */
    public static boolean isNotEmpty(String text) {
        return !isEmpty(text);
    }

    /**
     * 将字符串转为Unicode
     *
     * @param text 原始字符串
     * @return Unicode
     */
    public static String toUnicode(String text) {
        char[] chars = text.toCharArray();
        StringBuilder sb = new StringBuilder();
        for (char ch : chars) {
            sb.append("\\u").append(Integer.toString(ch, 16));
        }
        return sb.toString();
    }

    /**
     * 字符串是否是数字
     *
     * @param text 字符串
     * @return 是否是数字
     */
    public static boolean isNumber(String text) {
        Pattern pattern = Pattern.compile("^[\\d]+(\\.[\\d]+)?$");
        return pattern.matcher(text).matches();
    }

    /**
     * 是否是email
     *
     * @param text 字符串
     * @return 是否email
     */
    public static boolean isEmail(String text) {
        Pattern pattern = Pattern.compile("^(mailto:)?[A-Za-z0-9\\u4e00-\\u9fa5]+@[a-zA-Z0-9_-]+(\\.[a-zA-Z0-9_-]+)+$");
        return pattern.matcher(text).matches();
    }

    /**
     * 是否是文件地址
     *
     * @param text 字符串
     * @return 是否文件地址
     */
    public static boolean isFile(String text) {
        // false for moment.
        return false;
    }

    /**
     * 是否是url
     *
     * @param text 字符串
     * @return 是否url
     */
    private static boolean isUrl(String text) {
        return text.startsWith("https://") || text.startsWith("http://") || text.startsWith("ws://") || text.startsWith("file://") || text.startsWith("ftp://");
    }

    /**
     * 是否是超链接
     *
     * @param text 字符串
     * @return 是否超链接
     */
    public static boolean isHyperLink(String text) {
        return isEmail(text) || isUrl(text) || isFile(text);
    }

    /**
     * 获取字符串超链接类型
     *
     * @param text 字符串
     * @return 超链接类型
     */
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
