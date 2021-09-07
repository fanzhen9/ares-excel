package com.ares.excel.annotation;

import com.ares.excel.service.impl.DefaultDownLoadImageService;

import java.lang.annotation.*;

/**
 * @author fox
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {

    /**
     * 文件索引序号
     * @return
     */
    int index();

    /**
     * 文件中文名称
     * @return
     */
    String name();

    /**
     * 字段是否URL
     * @return
     */
    boolean isUrl() default false;

    /**
     * 字段类型
     * @return
     */
    Class fieldClassType() default String.class;

    /**
     * 图片下载类
     * @return
     */
    Class DownLoadImageService() default DefaultDownLoadImageService.class;
}
