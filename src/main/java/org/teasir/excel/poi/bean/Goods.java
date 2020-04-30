package org.teasir.excel.poi.bean;

import java.math.BigDecimal;


public class Goods {
    //商品编码
    private String code;
    //商品名字
    private String name;
    //商品数量
    private BigDecimal amount;
    //商品生产日期
    private String day;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }


    public BigDecimal getAmount() {
        return amount;
    }

    public void setAmount(BigDecimal amount) {
        this.amount = amount;
    }

    public String getDay() {
        return day;
    }

    public void setDay(String day) {
        this.day = day;
    }
}