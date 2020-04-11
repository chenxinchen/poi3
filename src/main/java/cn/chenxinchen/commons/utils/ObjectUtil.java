package cn.chenxinchen.commons.utils;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public class ObjectUtil {
    public static boolean isAllEmpty(Object o) {
        if (o == null) {
            return true;
        }
        Field[] fields = o.getClass().getDeclaredFields();
        int flag = 0;
        int length = fields.length;
        for (Field field : fields) {
            try {
                field.setAccessible(true);
                String typeName = field.getType().getName();
                Object resultValue = field.get(o);
                if (resultValue == null) {
                    flag++;
                } else {
                    if ("byte".equals(typeName) && (byte) resultValue == 0) {
                        flag++;
                    }
                    if ("short".equals(typeName) && (short) resultValue == 0) {
                        flag++;
                    }
                    if ("int".equals(typeName) && (int) resultValue == 0) {
                        flag++;
                    }
                    if ("long".equals(typeName) && (long) resultValue == 0) {
                        flag++;
                    }
                    if ("float".equals(typeName) && (float) resultValue == 0.0) {
                        flag++;
                    }
                    if ("double".equals(typeName) && (double) resultValue == 0.0) {
                        flag++;
                    }
                    if ("char".equals(typeName) && (char) resultValue == 0) {
                        flag++;
                    }
                    if ("boolean".equals(typeName) && !(boolean)resultValue) {
                        flag++;
                    }
                }
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
        return flag == length;
    }
}
