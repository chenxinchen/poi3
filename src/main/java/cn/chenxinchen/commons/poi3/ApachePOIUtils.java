package cn.chenxinchen.commons.poi3;

/**
 * ░░░░░░░░░░░░░░░░░░░░░░░░▄░░
 * ░░░░░░░░░▐█░░░░░░░░░░░▄▀▒▌░
 * ░░░░░░░░▐▀▒█░░░░░░░░▄▀▒▒▒▐
 * ░░░░░░░▐▄▀▒▒▀▀▀▀▄▄▄▀▒▒▒▒▒▐
 * ░░░░░▄▄▀▒░▒▒▒▒▒▒▒▒▒█▒▒▄█▒▐
 * ░░░▄▀▒▒▒░░░▒▒▒░░░▒▒▒▀██▀▒▌
 * ░░▐▒▒▒▄▄▒▒▒▒░░░▒▒▒▒▒▒▒▀▄▒▒
 * ░░▌░░▌█▀▒▒▒▒▒▄▀█▄▒▒▒▒▒▒▒█▒▐
 * ░▐░░░▒▒▒▒▒▒▒▒▌██▀▒▒░░░▒▒▒▀▄
 * ░▌░▒▄██▄▒▒▒▒▒▒▒▒▒░░░░░░▒▒▒▒
 * ▀▒▀▐▄█▄█▌▄░▀▒▒░░░░░░░░░░▒▒▒
 * 千万不要让你的小孩入这行
 */

import lombok.extern.slf4j.Slf4j;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 该工具类用于Excel和Word处理
 * <p>
 * HSSF － 提供读写Microsoft Excel格式档案的功能。
 * HWPF － 提供读写Microsoft Word格式档案的功能。
 *
 * @Author chenxinchen
 * @Version 1.0 ,POI版本 3.17,poi-ooxml 3.17
 * 用世界上最好的工具类，写最简单的业务代码！！！
 */
@Slf4j
public class ApachePOIUtils<T> {
    /**
     * 生成对应版本的Workbook
     *
     * @param file
     * @return
     * @throws IOException
     */
    private static Workbook getWorkbook(File file) throws IOException {
        // 1.获取文件名后缀
        String suffixName = null;
        try {
            suffixName = file.getName().substring(file.getName().lastIndexOf('.'));
        } catch (StringIndexOutOfBoundsException e) {
            log.error("该文件没有后缀，错误信息 {}", e);
            return null;
        }
        // 2.生成对应版本Workbook对象,07版.xlsx 03版.xls
        if (".xlsx".equals(suffixName))
            return new XSSFWorkbook(new FileInputStream(file));
        else if (".xls".equals(suffixName))
            return new HSSFWorkbook(new FileInputStream(file));
        else
            throw new RuntimeException("文件不是Excel文档！");
    }

    /**
     * 使用说明：
     * <br>
     * 读取文件中一个工作表，可指定行和列确定读取的位置
     * 你只要明确实体类，那接下来你就会得到封装好的List
     * <br>
     * 默认尽量把Excel内容解析到对象上(有风险，给你们加bug，手动狗头)
     *
     * @param file 文件对象
     * @param oClass 字节码对象
     * @param indexSheet 工作表
     * @param indexRowNum 第几行开始读
     * @param indexCellNum 第几列开始解析
     * @param <T> 泛型
     * @return
     * @throws IOException
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public static <T> List<T> readExcel(File file, Class<T> oClass, int indexSheet, int indexRowNum, int indexCellNum) throws IOException, IllegalAccessException, InstantiationException, NoSuchMethodException, InvocationTargetException {
        List<T> container = new ArrayList<>();
        // 1.获取Workbook对象
        Workbook workbook = getWorkbook(file);
        // 2.指定工作表
        Sheet sheet = workbook.getSheetAt(indexSheet);
        // 3.获取最大行数,有几行,excel中有1，2，3，4，5来表示行数,对应到程序中一行就是一个实例对象
        int rowNum = sheet.getLastRowNum() + 1;
        // 4.获取属性数组
        Field[] fields = oClass.getDeclaredFields();
        for (int i = indexRowNum; i < rowNum; i++) {
            // 5.准备好对象实例
            T o = oClass.newInstance();
            Row row = sheet.getRow(i);
            // 6.获取最大列数，有几列,excel中有A,B,C,D,E来表示列数,对应到程序中一列就是一个属性
            int cellNum = row.getLastCellNum();
            // 7.判断属性和列数数量相同
            if ((cellNum - indexCellNum) != fields.length)
                throw new RuntimeException("表格列数和属性数量不相同");
            int fieldIndex = 0;
            for (int j = indexCellNum; j < cellNum; j++) {
                // 8.获取每个单元格，啥都不填获取cell为null
                Cell cell = row.getCell(j);
                if (cell == null)
                    continue;
                // 9.获取属性名字，到时候调用属性的set方法来赋值
                Field field = fields[fieldIndex++];

                // 根据属性类型把内容强制赋值(尽可能把内容值转换属性类型，但会出现不可预知错误)
                assembly(oClass, o, cell, field);

                /*
                // 10.获取单元格类型,并尝试按顺序给实体类属性赋值,根据单元格内容类型给属性封装
                // (可能导致内容有值，属性却没值,但能保证没有错误，且不会出现数据任何异常)
                switch (cell.getCellTypeEnum()) {
                    case _NONE:
                        log.debug("未知类型");
                        break;
                    case NUMERIC:
                        log.debug("数字类型，整数、小数、日期");
//                        log.info(cell.getNumericCellValue() + "");
                        assembly(oClass, o, cell, field);
                        break;
                    case STRING:
                        log.debug("文本类型，等同String");
//                        log.info(cell.getStringCellValue());
                        assembly(oClass, o, cell, field);
                        break;
                    case FORMULA:
                        log.debug("公式类型，没理解什么类型");
                        break;
                    case BLANK:
                        log.debug("空白类型，现在不确定空白是空字符串，还是null");
                        break;
                    case BOOLEAN:
                        log.debug("布尔类型，等同boolean");
//                        log.info(cell.getBooleanCellValue() + "");
                        assembly(oClass, o, cell, field);
                        break;
                    case ERROR:
                        log.debug("错误类型，不知道干嘛用的");
                        log.info(cell.getErrorCellValue() + "");
                        break;
                }
                */
            }
            container.add(o);
        }
        // ?.收尾工作
        workbook.close();
        return container;
    }

    /**
     * 专门用于封装对象
     *
     * @param oClass 字节码对象
     * @param o 实例对象
     * @param cell 单元格
     * @param field 属性
     * @param <T> 泛型
     * @throws NoSuchMethodException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private static <T> void assembly(Class<T> oClass, T o, Cell cell, Field field) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        String fieldNameUpperCase = field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);
        String typeName = field.getGenericType().getTypeName();
        try {
            if ("byte".equals(typeName) || "java.lang.Byte".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "byte".equals(typeName) ? byte.class : Byte.class);
                method.invoke(o, (byte) cell.getNumericCellValue());
            }
            if ("short".equals(typeName) || "java.lang.Short".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "short".equals(typeName) ? short.class : Short.class);
                method.invoke(o, (short) cell.getNumericCellValue());
            }
            if ("int".equals(typeName) || "java.lang.Integer".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "int".equals(typeName) ? int.class : Integer.class);
                method.invoke(o, (int) cell.getNumericCellValue());
            }
            if ("long".equals(typeName) || "java.lang.Long".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "long".equals(typeName) ? long.class : Long.class);
                method.invoke(o, (long) cell.getNumericCellValue());
            }
            if ("float".equals(typeName) || "java.lang.Float".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "float".equals(typeName) ? float.class : Float.class);
                method.invoke(o, (float) cell.getNumericCellValue());
            }
            if ("double".equals(typeName) || "java.lang.Double".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "double".equals(typeName) ? double.class : Double.class);
                method.invoke(o, cell.getNumericCellValue());
            }
            if ("char".equals(typeName) || "java.lang.Character".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "char".equals(typeName) ? char.class : Character.class);
                method.invoke(o, cell.getStringCellValue().charAt(0));
            }
            if ("boolean".equals(typeName) || "java.lang.Boolean".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "boolean".equals(typeName) ? boolean.class : Boolean.class);
                method.invoke(o, cell.getBooleanCellValue());
            }
            if ("java.lang.String".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, String.class);
                method.invoke(o, cell.getStringCellValue());
            }
            if ("java.util.Date".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, Date.class);
                method.invoke(o, new Date((long) cell.getNumericCellValue()));
            }
        }catch (IllegalStateException e){
            if ("java.lang.String".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, String.class);
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        method.invoke(o, new Double(cell.getNumericCellValue()).toString());
                        break;
                    case BOOLEAN:
                        method.invoke(o, new Boolean(cell.getBooleanCellValue()).toString());
                        break;
                }
            }else if ("char".equals(typeName) || "java.lang.Character".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "char".equals(typeName) ? char.class : Character.class);
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        method.invoke(o, new Double(cell.getNumericCellValue()).toString().charAt(0));
                        break;
                    case BOOLEAN:
                        method.invoke(o, new Boolean(cell.getBooleanCellValue()).toString().charAt(0));
                        break;
                }
            }else
                log.error("单元格内容转换不了属性类型 {}",e);
        }
    }
}
