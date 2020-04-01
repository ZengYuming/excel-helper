package com.tony.demo.excel_helper.extension;

import com.tony.demo.excel_helper.exception.ExcelContentInvalidException;

/**
 * 字段校验规则
 * Created by TonyZeng on 2017/9/4.
 */
public interface ExcelRule<T> {
    /**
     * 检查
     *
     * @param value
     * @throws ExcelContentInvalidException
     */
    public void check(Object value, String columnName, String fieldName) throws ExcelContentInvalidException;
}