package com.yhy.doc.excel.entity;

import com.yhy.doc.excel.CategoryFilter;
import com.yhy.doc.excel.SexConverter;
import com.yhy.doc.excel.annotation.*;
import com.yhy.doc.excel.offer.DateParser;
import com.yhy.doc.excel.offer.LocalDateTimeParser;
import com.yhy.doc.excel.offer.TimestampParser;
import lombok.Data;
import lombok.ToString;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.io.Serializable;
import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.Random;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-10 0:27
 * version: 1.0.0
 * desc   :
 */
@Data
@ToString
@Document
public class Company implements Serializable {
    private static final long serialVersionUID = 198093835962499501L;
    private static final Random rand = new Random();

    @Excel("序号")
    @Align(
            horizontal = HorizontalAlignment.CENTER,
            vertical = VerticalAlignment.CENTER
    )
    private int index;

    @Excel("创建日期")
    @Border
    @Parser(DateParser.class)
    private Date createDate;

    @Excel("更新日期")
    @Border
    @Parser(TimestampParser.class)
    private Timestamp updateDate;

    @Excel("删除日期")
    @Border
    @Parser(LocalDateTimeParser.class)
    private LocalDateTime deleteDate;

    @Excel(value = "商户名称", wrap = true)
    @Ground(
            back = IndexedColors.AQUA,
            fore = IndexedColors.AQUA,
            pattern = FillPatternType.SOLID_FOREGROUND
    )
    @Border(color = IndexedColors.DARK_RED)
    private String name;

    @Excel("法人性别")
    @Sorted(2)
    @Converter(SexConverter.class)
    @Font(size = 8)
    private Sex sex;

    @Excel("商户类型")
    @Font(underline = org.apache.poi.ss.usermodel.Font.U_DOUBLE)
    private String type;

    @Excel("联系人")
    @Font(delete = true)
    private String contact;

    @Excel("联系电话")
    @Border(color = IndexedColors.TURQUOISE)
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
    @Style(
            align = @Align(
                    horizontal = HorizontalAlignment.CENTER,
                    vertical = VerticalAlignment.CENTER
            ),
            font = @Font(size = 18)
    )
    private String code;

    @Excel("法人代表")
    @Ground
    private String law;

    @Excel(value = "法人代表证件类型")
    @Ground(
            fore = IndexedColors.YELLOW,
            back = IndexedColors.AQUA,
            pattern = FillPatternType.SQUARES
    )
    @Font(color = IndexedColors.RED)
    @Border
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
    private float percent = 0.12F;

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

    @Excel("加数1")
    private int addA = rand.nextInt(101);

    @Excel("加数2")
    private int addB = rand.nextInt(101);

    @Excel(value = "和", formula = "SUM(AE{},AF{})")
    private int sum;
}
