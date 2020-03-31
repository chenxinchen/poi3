package cn.chenxinchen.commons;

import cn.chenxinchen.commons.poi3.ApachePOIUtils;
import cn.chenxinchen.commons.pojo.Goods;
import cn.chenxinchen.commons.pojo.User;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

public class PoiToolTest {
    @Test
    public void readExcelText() throws InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, IOException {
        List<User> users = ApachePOIUtils.readExcel(new File("C:\\Users\\chenxinchen\\Desktop\\我工具类测试数据.xlsx"), User.class, 0, 3, 0);
        System.out.println(users);
    }
    @Test
    public void readExcelUseAnnoText() throws InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, IOException {
        List<Goods> goods = ApachePOIUtils.readExcel(new File("C:\\Users\\chenxinchen\\Desktop\\我工具类测试数据.xlsx"), Goods.class, 1, 3, 0);
        System.out.println(goods);
    }
}
