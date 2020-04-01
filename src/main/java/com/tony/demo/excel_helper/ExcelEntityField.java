package com.tony.demo.excel_helper;

import com.tony.demo.excel_helper.annotation.ExcelProperty;

import java.lang.reflect.Field;

/**
 * Excel实体字段类（内部类）
 * Created by TonyZeng on 2017/9/6.
 */
class ExcelEntityField {
    private String columnName;
    private boolean required;
    private Field field;
    private int index;
    private ExcelProperty annotation;

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public ExcelProperty getAnnotation() {
        return annotation;
    }

    public void setAnnotation(ExcelProperty annotation) {
        this.annotation = annotation;
    }
}