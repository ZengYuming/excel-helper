package com.tony.demo.excel_helper.annotation;


import com.tony.demo.excel_helper.extension.ExcelRule;
import com.tony.demo.excel_helper.rule.NoneRule;

import java.lang.annotation.*;

/**
 * 标记字段为Excel填充字段
 * Created by TonyZeng on 2017/9/4.
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelProperty {
    /**
     * excel的列名
     */
    String value() default "";

    /**
     * 是否必填
     */
    boolean required() default false;

    /**
     * 校验规则类
     */
    @SuppressWarnings("rawtypes") Class<? extends ExcelRule> rule() default NoneRule.class;

    /**
     * 正则表达式校验规则
     * 该项仅对String, Long, Short,Integer类型有效
     */
    String regexp() default "";

    /**
     * 正则规则校验失败错误提示信息
     */
    String regexpErrorMessage() default "";

    /**
     * 默认值
     * 默认值均采用String类型，系统将会自动进行类型转换，不支持对象类型!
     */
    String defaultValue() default "";

    /**
     * 是否采用默认值
     * 如果为true，默认值为空的时候会使用空字符串，否则使用null
     */
    boolean hasDefaultValue() default false;
}
