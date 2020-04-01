package com.tony.demo.excel_helper.consts;

/**
 * Excel类型
 * Created by TONZENG on 2017-06-19.
 */
public enum ExcelType {
    xlsx(1, "xslx"),//2007及以后版本  .xslx
    xls(2, "xsl");//2003及以前的版本 .xsl

    private int id;
    private String chinese;

    // 构造方法
    ExcelType(int id, String chinese) {
        this.id = id;
        this.chinese = chinese;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getChinese() {
        return chinese;
    }

    public void setChinese(String chinese) {
        this.chinese = chinese;
    }

    public static String getChinese(int id) {
        for (ExcelType a : ExcelType.values()) {
            if (a.getId() == id) {
                return a.chinese;
            }
        }
        return "";
    }

    public static Integer getId(String chinese) {
        for (ExcelType a : ExcelType.values()) {
            if (a.getChinese().equals(chinese)) {
                return a.id;
            }
        }
        return 0;
    }
}
