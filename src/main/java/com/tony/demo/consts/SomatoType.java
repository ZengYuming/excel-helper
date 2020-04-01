package com.tony.demo.consts;

/**
 * 体型
 * Created by TONZENG on 2017-06-19.
 */
public enum SomatoType {
    Wasting(1, "消瘦"),
    Standard(2, "标准"),
    RecessiveObesity(3, "隐性肥胖"),
    Robust(4, "健壮"),
    Obesity(5, "肥胖");

    private int id;
    private String chinese;

    // 构造方法
    SomatoType(int id, String chinese) {
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
        for (SomatoType a : SomatoType.values()) {
            if (a.getId() == id) {
                return a.chinese;
            }
        }
        return "";
    }
    public static Integer getId(String chinese) {
        for (SomatoType a : SomatoType.values()) {
            if (a.getChinese().equals(chinese)) {
                return a.id;
            }
        }
        return 0;
    }
}
