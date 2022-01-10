package com.wecool.poiexport;

import java.lang.annotation.*;

/**
 * @author bowafterrain [mazhaoming@vip.qq.com]
 * @date 2022-1-6 22:32
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Repeatable(ExcelColumn.List.class)
public @interface ExcelColumn {

    String name();

    int order();

    String datePattern() default "yyyy-MM-dd HH:mm:ss";

    Class<?> group() default DefaultGroup.class;

    interface DefaultGroup {
    }

    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    @Documented
    @interface List {
        ExcelColumn[] value();
    }
}
