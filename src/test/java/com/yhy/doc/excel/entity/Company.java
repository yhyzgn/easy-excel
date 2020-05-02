package com.yhy.doc.excel.entity;

import com.yhy.doc.excel.CategoryFilter;
import com.yhy.doc.excel.SexConverter;
import com.yhy.doc.excel.annotation.*;
import com.yhy.doc.excel.offer.DateParser;
import lombok.Data;
import lombok.ToString;

import java.io.Serializable;
import java.util.Date;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 0:27
 * version: 1.0.0
 * desc   :
 */
@Data
@ToString
public class Company implements Serializable {
    private static final long serialVersionUID = 198093835962499501L;

    @Excel("序号")
    private int index;

    @Excel("创建日期")
    @Parser(DateParser.class)
    private Date createDate;

    @Excel(value = "商户名称", wrap = true)
    private String name;

    @Excel("法人性别")
    @Sorted(2)
    @Converter(SexConverter.class)
    private Sex sex;

    @Excel("商户类型")
    private String type;

    @Excel("联系人")
    private String contact;

    @Excel("联系电话")
    private String phone;

    @Excel("联系邮箱")
    private String email;

    @Excel("所属行业")
    @Filter(CategoryFilter.class)
    private String category;

    @Excel(value = "门店名称", wrap = true)
    private String storeName;

    @Excel("所在州市")
    @Ignored
    private String city;

    @Excel("所在区县")
    private String county;

    @Excel(like = "%地址", wrap = true)
    private String address;

    @Excel(value = "统一信用代码", nullable = false, wrap = true, tolerance = 0.8)
    @Sorted(1)
    private String code;

    @Excel("法人代表")
    private String law;

    @Excel(value = "法人代表证件类型")
    private String cardType;

    @Excel("法人证件号")
    private String cardNumber;

    @Excel("string")
    private String str = "12";

    @Excel("int")
    private int integer = 12;

    @Excel("zhCN")
    @Pattern("[DbNum2][$-804]0")
    private int zhCN = 12;

    @Excel("float")
    private float flot = 12;

    @Excel("percent")
    @Pattern("0.00%")
    private float percent = 1.02F;

    @Excel("byte")
    private byte bt = 12;

    @Excel("boolean")
    private boolean bln = true;

    @Excel("long")
    private long lng = 122342342342L;

    @Excel("short")
    private short shot = 12;

    @Excel("double")
    private double dobl = 122342342342.034400D;

    @Excel("money")
    @Pattern("￥#,##0")
    private double money = 122342342342.034400D;

    @Excel("char")
    private char ch = 12;
}
