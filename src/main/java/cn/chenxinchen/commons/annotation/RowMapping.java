package cn.chenxinchen.commons.annotation;

import java.lang.annotation.*;

/**
 * 标志类上，用于判断是否使用注解解析，<br>
 * 并可以指定从哪行开始解析,默认第1行开始解析
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface RowMapping {
    int value() default 1;
}
