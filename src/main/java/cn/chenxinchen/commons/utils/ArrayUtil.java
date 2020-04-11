package cn.chenxinchen.commons.utils;

public class ArrayUtil {
    public static boolean isAllEmpty(String[] arrays) {
        if(arrays == null || arrays.length == 0){
            return true;
        }
        int length = arrays.length;
        int flag = 0;
        for (String array : arrays) {
            if(array == null || "".equals(array.trim())){
                flag++;
            }
        }
        return flag == length;
    }
}
