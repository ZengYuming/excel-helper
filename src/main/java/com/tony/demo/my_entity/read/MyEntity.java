package com.tony.demo.my_entity.read;


import com.tony.demo.my_checkRule.MyCheckRule;
import com.tony.demo.excel_helper.annotation.ExcelEntity;
import com.tony.demo.excel_helper.annotation.ExcelProperty;

import java.sql.Timestamp;

/**
 * Created by TonyZeng on 2017/9/4.
 */
@ExcelEntity
public class MyEntity {
    @ExcelProperty(value = "姓名", rule = MyCheckRule.class)
    private String name;
    @ExcelProperty(value = "性别", required = true)
    private String sex;
    @ExcelProperty(value = "年龄", regexp = "^[1-4]{1}[0-9]{1}$", regexpErrorMessage = "年龄必须在10-49岁之间")
    private Integer age;
    @ExcelProperty("体型")
    private MyFieldType somato;
    @ExcelProperty(value = "电话")
    private Long tel;
    @ExcelProperty("创建时间")
    private Timestamp createDate;


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public MyFieldType getSomato() {
        return somato;
    }

    public void setSomato(MyFieldType somato) {
        this.somato = somato;
    }

    public Timestamp getCreateDate() {
        return createDate;
    }

    public void setCreateDate(Timestamp createDate) {
        this.createDate = createDate;
    }

    public Long getTel() {
        return tel;
    }

    public void setTel(Long tel) {
        this.tel = tel;
    }

    @Override
    public String toString() {
        return "MyEntity{" +
                "name='" + name + '\'' +
                ", sex='" + sex + '\'' +
                ", age=" + age +
                ", somato=" + somato +
                ", tel=" + tel +
                ", createDate=" + createDate +
                '}';
    }
}
