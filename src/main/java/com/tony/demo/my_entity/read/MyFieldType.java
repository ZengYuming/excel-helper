package com.tony.demo.my_entity.read;

import com.tony.demo.consts.SomatoType;
import com.tony.demo.excel_helper.annotation.ExcelProperty;
import com.tony.demo.excel_helper.extension.ExcelFieldType;

import java.lang.reflect.Field;

/**
 * 自定义类型-体型
 * <p>
 * Created by TonyZeng on 2017/9/5.
 */
public class MyFieldType extends ExcelFieldType {
    public MyFieldType() {
    }

    public MyFieldType(String somato) {
        this.somato = somato;
    }

    public String somato;

    @Override
    public Object parseValue(String value) {
        MyFieldType somatoFieldType = new MyFieldType();
        try {
            somatoFieldType.setSomato(SomatoType.getChinese(Integer.parseInt(value)));
        } catch (Exception e) {
            return null;
        }
        return somatoFieldType;
    }

    public String getSomato() {
        return somato;
    }

    public void setSomato(String somato) {
        this.somato = somato;
    }

    @Override
    public String toString() {
        return somato;
    }

    /**
     * Excel实体字段类（内部类）
     */
    private class ExcelEntityField {
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


}