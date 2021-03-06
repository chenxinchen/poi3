package cn.chenxinchen.commons.poi3;

import cn.chenxinchen.commons.annotation.RowMapping;
import cn.chenxinchen.commons.pojo.Goods;
import org.junit.Test;

public class MyAnnotationTest {
    @Test
    public void hasAnnotation() {
        Class<? extends Goods> aClass = Goods.class;
        RowMapping rowMapping = aClass.getDeclaredAnnotation(RowMapping.class);
        System.out.println(rowMapping.value());
        /*Field[] declaredFields = aClass.getDeclaredFields();
        for (Field declaredField : declaredFields) {
            ColumnMapping declaredAnnotation = declaredField.getDeclaredAnnotation(ColumnMapping.class);
            ColumnSerial value = declaredAnnotation.value();
            System.out.println(value.getIndex());
        }*/
    }
}
