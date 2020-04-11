package cn.chenxinchen.commons.annotation;

import java.lang.annotation.*;

/**
 * 用于标志属性在表格哪个位置
 * <p>
 * 作者: 陈信晨
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ColumnMapping {
    ColumnSerial value();
}
