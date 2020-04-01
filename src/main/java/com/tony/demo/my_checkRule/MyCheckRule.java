package com.tony.demo.my_checkRule;


import com.tony.demo.excel_helper.exception.ExcelContentInvalidException;
import com.tony.demo.excel_helper.extension.ExcelRule;

/**
 * Created by TonyZeng on 2017/9/4.
 */
public class MyCheckRule implements ExcelRule<String> {
    // 字段检查
    public void check(Object value, String columnName, String fieldName) throws ExcelContentInvalidException {
        String val = (String) value;
        if (val.length() > 10) {
            throw new ExcelContentInvalidException("内容超长!");
        }
    }
}