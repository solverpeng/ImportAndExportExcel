package com.solverpeng.poi.servlet;

import com.solverpeng.poi.beans.User;
import com.solverpeng.poi.utils.excel.ExportExcelUtil;
import org.apache.commons.collections.CollectionUtils;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * Created by solverpeng on 2017/2/17 0017.
 */
public class PoiServlet2 extends HttpServlet {
    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        this.doPost(req, resp);
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        User user1 = new User("1", "tom", 23, new Date(), "北京");
        User user2 = new User("2", "jerry", 25, new Date(), "上海");
        User user3 = new User("3", "lily", 26, new Date(), "洛杉矶");
        User user4 = new User("4", "lucy", 27, new Date(), "纽约");
        List<User> list = Arrays.asList(user1, user2, user3, user4);

        List<Object[]> objects = new ArrayList<>();
        if (CollectionUtils.isNotEmpty(list)) {
            for (User user : list) {
                Object[] objectArr = new Object[4];
                objectArr[0] = user.getUserName();
                objectArr[1] = user.getAge();
                objectArr[2] = user.getBirth();
                objectArr[3] = user.getAddress();
                objects.add(objectArr);
            }
        }

        String titleName = "用户信息";
        List<String> headers = Arrays.asList("用户姓名", "用户年龄", "用户生日", "居住地址");
        ExportExcelUtil.downLoad4Excel(titleName, headers, objects, "用户信息", req, resp);
    }
}