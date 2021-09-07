package com.ares.excel.annotation;

import java.lang.annotation.*;

/**
 * @author fox
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelSheet {


    /**
     * sheet 索引
     * @return
     */
    int index();

    /**
     * sheet 名称
     * @return
     */
    String name();
}
