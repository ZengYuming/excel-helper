package com.tony.demo.excel_helper.annotation;

import java.lang.annotation.*;

/**
 * 标记实体为Excel实体
 * Created by TonyZeng on 2017/9/4.
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelEntity {
}
