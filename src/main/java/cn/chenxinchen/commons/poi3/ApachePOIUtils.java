package cn.chenxinchen.commons.poi3;

import cn.chenxinchen.commons.annotation.ColumnMapping;
import cn.chenxinchen.commons.annotation.ColumnSerial;
import cn.chenxinchen.commons.annotation.RowMapping;
import cn.chenxinchen.commons.utils.ArrayUtil;
import cn.chenxinchen.commons.utils.ObjectUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 该工具类用于Excel和Word处理
 * <p>
 * 作者:陈信晨
 * 版本: 1.0
 * 用世界上最好的工具类，写最简单的业务代码！！！
 */
@Slf4j
public class ApachePOIUtils {
    /**
     * 生成对应版本的Workbook
     *
     * @param file 文件对象
     * @return Workbook
     * @throws IOException IO异常
     */
    private static Workbook getWorkbook(File file) throws IOException {
        // 获取文件名后缀
        String suffixName = file.getName().substring(file.getName().lastIndexOf('.'));
        // 生成对应版本Workbook对象,07版.xlsx 03版.xls
        if (".xlsx".equals(suffixName))
            return new XSSFWorkbook(new FileInputStream(file));
        else if (".xls".equals(suffixName))
            return new HSSFWorkbook(new FileInputStream(file));
        else
            throw new RuntimeException("文件不是Excel文档！");
    }

    private static Workbook getWorkbook(InputStream is, String fileName) throws IOException {
        // 获取文件名后缀
        String suffixName = fileName.substring(fileName.lastIndexOf('.'));
        // 生成对应版本Workbook对象,07版.xlsx 03版.xls
        if (".xlsx".equals(suffixName))
            return new XSSFWorkbook(is);
        else if (".xls".equals(suffixName))
            return new HSSFWorkbook(is);
        else
            throw new RuntimeException("文件不是Excel文档！");
    }

    /**
     * 使用说明：
     * <p>
     * 读取文件中一个工作表，可指定行和列确定读取的位置
     * 你只要明确实体类，那接下来你就会得到封装好的List
     * <p>
     * 默认尽量把Excel内容解析到对象上(有风险，给你们加bug，手动狗头)
     * <p>
     * 当实体类上有注解时，参数indexRowNum和indexCellNum不生效
     *
     * @param workbook    Excel对象
     * @param oClass      字节码对象
     * @param indexSheet  工作表
     * @param indexRowNum 以Excel行数为准，第几行解析
     * @param indexCell   以Excel列数为准，第几列解析
     * @param <T>         泛型
     * @return 一个List对象
     */
    private static <T> List<T> readExcel(Workbook workbook, Class<T> oClass, int indexSheet, Integer indexRowNum, ColumnSerial indexCell) throws IOException, IllegalAccessException, InstantiationException, NoSuchMethodException, InvocationTargetException {
        List<T> container = new ArrayList<>();
        // 指定工作表
        Sheet sheet = workbook.getSheetAt(indexSheet);
        // 获取Excel表格最大行数
        int rowNum = sheet.getLastRowNum() + 1;
        // 添加解析策略，当实体类有RowMapping注释使用注释解析，没有注释使用默认解析
        RowMapping rowMapping = oClass.getDeclaredAnnotation(RowMapping.class);
        // 准备好对象实例
        T o = oClass.newInstance();
        if (rowMapping == null) {
            // 默认解析按属性顺序和表格顺序解析过去
            Field[] fields = oClass.getDeclaredFields();
            for (int i = indexRowNum - 1; i < rowNum; i++) {
                Row row = sheet.getRow(i);
                // 获取最大列数
                int cellNum = row.getLastCellNum();
                // 判断属性和列数数量相同
                if ((cellNum - indexCell.getIndex()) > fields.length)
                    throw new RuntimeException("表格列数和属性不兼容，表格列数:" + (cellNum - indexCell.getIndex()) + "属性个数:" + fields.length);
                int fieldIndex = 0;
                for (int j = indexCell.getIndex(); j < cellNum; j++) {
                    // 获取每个单元格，啥都不填获取cell为null
                    Cell cell = row.getCell(j);
                    if (cell == null) {
                        fieldIndex++;
                    } else {
                        // 获取属性名字，到时候调用属性的set方法来赋值
                        Field field = fields[fieldIndex++];
                        // 根据属性类型把内容强制赋值(尽可能把内容值转换属性类型，但会出现不可预知错误)
                        assembly(oClass, o, cell, field);
                    }

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
            }
        } else {
            // 使用注解解析
            Field[] fields = oClass.getDeclaredFields();
            for (int i = rowMapping.value() - 1; i < rowNum; i++) {
                // 准备表格行对象
                Row row = sheet.getRow(i);
                // 循环实体类属性
                for (Field field : fields) {
                    ColumnMapping columnMapping = field.getDeclaredAnnotation(ColumnMapping.class);
                    if (columnMapping != null) {
                        // 获取解析列下标
                        int index = columnMapping.value().getIndex();
                        assembly(oClass, o, row.getCell(index), field);
                    }
                }
            }
            if (!ObjectUtil.isAllEmpty(o)) {
                container.add(o);
            }
        }
        // 关闭资源
        workbook.close();
        return container;
    }
    @SuppressWarnings("all")
    public static <T> List<T> readExcel(File file, Class<T> oClass, int indexSheet, Integer indexRowNum, ColumnSerial indexCell) throws Exception {
        return readExcel(getWorkbook(file), oClass, indexSheet, indexRowNum, indexCell);
    }

    public static <T> List<T> readExcel(InputStream is, String fileName, Class<T> oClass, int indexSheet, Integer indexRowNum, ColumnSerial indexCell) throws Exception {
        return readExcel(getWorkbook(is, fileName), oClass, indexSheet, indexRowNum, indexCell);
    }

    /**
     * 读取表格全部内容全以字符串存储
     *
     * @param workbook    Excel对象
     * @param indexSheet  工作表
     * @param indexRowNum 第几行开始解析
     * @param indexCell   第几列开始解析
     * @param endCell     第几列停止解析（传null按照数据实际列数解析） 可固定数组长度
     * @return List<String [ ]>
     */
    private static List<String[]> readExcel2String(Workbook workbook, int indexSheet, int indexRowNum, ColumnSerial indexCell, ColumnSerial endCell) throws IOException {
        // 准备容器
        List<String[]> container = new ArrayList<>();
        // 指定工作表
        Sheet sheet = workbook.getSheetAt(indexSheet);
        // 获取Excel表格最大行数
        int rowNum = sheet.getLastRowNum() + 1;
        // 循环行数
        for (int i = indexRowNum - 1; i < rowNum; i++) {
            // 获取行对象
            Row row = sheet.getRow(i);
            // 获取列最大列数
            short cellNum = row.getLastCellNum();
            if (endCell != null) {
                cellNum = (short) endCell.getIndex();
                cellNum++;
            }
            // 准备列的存放数据
            String[] data = new String[cellNum];
            int dataIndex = 0;
            for (int j = indexCell.getIndex(); j < cellNum; j++) {
                // 获取单元格
                Cell cell = row.getCell(j);
                if (cell == null) {
                    data[dataIndex] = null;
                } else {
                    // 判断数据并填充
                    switch (cell.getCellTypeEnum()) {
                        case _NONE:
                            log.debug("第" + (i + 1) + "行" + (j + 1) + "未知类型");
                            break;
                        case NUMERIC:
                            log.debug("第" + (i + 1) + "行" + (j + 1) + "数字类型，整数、小数、日期，value:" + cell.getNumericCellValue());
                            data[dataIndex] = new BigDecimal(cell.getNumericCellValue()).toString();
                            break;
                        case STRING:
                            log.debug("第" + (i + 1) + "行" + (j + 1) + "文本类型，value:" + cell.getStringCellValue());
                            data[dataIndex] = cell.getStringCellValue();
                            break;
                        case FORMULA:
                            log.debug("第" + (i + 1) + "行" + (j + 1) + "公式类型");
                            try {
                                data[dataIndex] = String.valueOf(cell.getNumericCellValue());
                            } catch (Exception e) {
                                data[dataIndex] = String.valueOf(cell.getStringCellValue());
                            }
                            break;
                        case BLANK:
                            log.debug("第" + (i + 1) + "行" + (j + 1) + "空类型,value:" + cell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            log.debug("第" + (i + 1) + "行" + (j + 1) + "布尔类型,value:" + cell.getBooleanCellValue());
                            data[dataIndex] = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case ERROR:
                            log.debug("第" + (i + 1) + "行" + (j + 1) + "错误类型,value:" + cell.getErrorCellValue());
                            break;
                    }
                }
                dataIndex++;
            }
            if (!ArrayUtil.isAllEmpty(data)) {
                container.add(data);
            }
        }
        workbook.close();
        return container;
    }
    @SuppressWarnings("all")
    public static List<String[]> readExcel2String(File file, int indexSheet, int indexRowNum, ColumnSerial indexCell, ColumnSerial endCell) throws Exception {
        return readExcel2String(getWorkbook(file), indexSheet, indexRowNum, indexCell, endCell);
    }

    public static List<String[]> readExcel2String(InputStream is, String fileName, int indexSheet, int indexRowNum, ColumnSerial indexCell, ColumnSerial endCell) throws Exception {
        return readExcel2String(getWorkbook(is, fileName), indexSheet, indexRowNum, indexCell, endCell);
    }

    /**
     * 专门用于封装对象
     *
     * @param oClass 字节码对象
     * @param o      实例对象
     * @param cell   单元格
     * @param field  属性
     * @param <T>    泛型
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
        } catch (IllegalStateException e) {
            if ("java.lang.String".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, String.class);
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        method.invoke(o, new BigDecimal(cell.getNumericCellValue()).toString());
                        break;
                    case BOOLEAN:
                        method.invoke(o, Boolean.toString(cell.getBooleanCellValue()));
                        break;
                }
            } else if ("char".equals(typeName) || "java.lang.Character".equals(typeName)) {
                Method method = oClass.getMethod("set" + fieldNameUpperCase, "char".equals(typeName) ? char.class : Character.class);
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        method.invoke(o, Double.toString(cell.getNumericCellValue()).charAt(0));
                        break;
                    case BOOLEAN:
                        method.invoke(o, Boolean.toString(cell.getBooleanCellValue()).charAt(0));
                        break;
                }
            } else {
                String cellString = null;
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        cellString = new BigDecimal(cell.getNumericCellValue()).toString();
                        break;
                    case STRING:
                        cellString = cell.getStringCellValue();
                        break;
                    case FORMULA:
                        try {
                            cellString = String.valueOf(cell.getNumericCellValue());
                        } catch (Exception ex) {
                            cellString = String.valueOf(cell.getStringCellValue());
                        }
                        break;
                    case BOOLEAN:
                        cellString = String.valueOf(cell.getBooleanCellValue());
                        break;
                }
                log.debug("单元格内容转换不了属性类型,单元格内容:" + cellString + ",属性类型:" + typeName, e);
            }

        }
    }
}
