package com.tony.demo.excel_helper.extension;

/**
 * 自定义字段类型
 * Created by TonyZeng on 2017/9/5.
 */
public abstract class ExcelFieldType<T> {
    public abstract T parseValue(String value);

    @Override
    public int hashCode() {
        return getClass().hashCode();
    }
}
