package cn.chenxinchen.commons;

import cn.chenxinchen.commons.annotation.ColumnSerial;
import cn.chenxinchen.commons.poi3.ApachePOIUtils;
import cn.chenxinchen.commons.pojo.Goods;
import cn.chenxinchen.commons.pojo.User;
import org.junit.Test;

import java.io.File;
import java.util.Arrays;
import java.util.List;

public class PoiToolTest {
    @Test
    public void readExcelTest() throws Exception {
        List<User> users = ApachePOIUtils.readExcel(new File("C:\\Users\\chenxinchen\\Desktop\\我工具类测试数据.xlsx"), User.class, 0, 5, ColumnSerial.A);
        System.out.println(users);
    }
    @Test
    public void readExcelUseAnnoTest() throws Exception {
        List<Goods> goods = ApachePOIUtils.readExcel(new File("C:\\Users\\chenxinchen\\Desktop\\我工具类测试数据.xlsx"), Goods.class, 1, 4, null);
        System.out.println(goods);
    }
    @Test
    public void readExcel2String() throws Exception {
        List<String[]> strings = ApachePOIUtils.readExcel2String(new File("C:\\Users\\chenxinchen\\Desktop\\我工具类测试数据.xlsx"), 0, 4, ColumnSerial.A);
        for (String[] string : strings) {
            System.out.println(Arrays.toString(string));
        }
    }
}
