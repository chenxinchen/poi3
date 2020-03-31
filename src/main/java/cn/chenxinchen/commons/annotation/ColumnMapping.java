package cn.chenxinchen.commons.annotation;

import java.lang.annotation.*;

/**
 * 用于标志属性在表格哪个位置
 *
 * 适用示例
 * @ColumnMapping(value = ColumnSerial.A)
 *
 * @Author chenxinchen
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ColumnMapping {
    ColumnSerial value();
}
