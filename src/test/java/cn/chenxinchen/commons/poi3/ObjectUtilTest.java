package cn.chenxinchen.commons.poi3;

import cn.chenxinchen.commons.pojo.User;
import cn.chenxinchen.commons.utils.ObjectUtil;
import org.junit.Test;

public class ObjectUtilTest {
    @Test
    public void test(){
        User u = new User();
        u.setUsername("21312");
        System.out.println(ObjectUtil.isAllEmpty(u));
        User u1 = new User();
        System.out.println(ObjectUtil.isAllEmpty(u1));
    }
}
