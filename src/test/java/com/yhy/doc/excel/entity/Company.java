package com.yhy.doc.excel.entity;

import com.yhy.doc.excel.SexConverter;
import com.yhy.doc.excel.annotation.Converter;
import com.yhy.doc.excel.annotation.Excel;
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
    private Date createDate;

    @Excel(value = "商户名称", wrap = true)
    private String name;

    @Excel("法人性别")
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
    private String category;

    @Excel(value = "门店名称", wrap = true)
    private String storeName;

    @Excel("所在州市")
    private String city;

    @Excel("所在区县")
    private String county;

    @Excel(like = "%地址", wrap = true)
    private String address;

    @Excel(value = "统一信用代码", nullable = false, wrap = true, tolerance = 0.8)
    private String code;

    @Excel("法人代表")
    private String law;

    @Excel(value = "法人代表证件类型")
    private String cardType;

    @Excel("法人证件号")
    private String cardNumber;
}
