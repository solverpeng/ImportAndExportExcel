package com.solverpeng.poi.servlet;

import com.solverpeng.poi.beans.User;
import com.solverpeng.poi.utils.excel.ExportExcel;
import org.apache.poi.ss.usermodel.Row;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * Created by solverpeng on 2017/2/16 0016.
 */
public class PoiServlet extends HttpServlet{

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        doPost(req, resp);
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        User user1 = new User("1", "tom", 23, new Date(), "北京");
        User user2 = new User("2", "jerry", 25, new Date(), "上海");
        User user3 = new User("3", "lily", 26, new Date(), "洛杉矶");
        User user4 = new User("4", "lucy", 27, new Date(), "纽约");
        List<User> list = Arrays.asList(user1, user2, user3, user4);

        // 导出模板
        /*ExportExcel exportExcel = new ExportExcel("员工信息模板", User.class, 2);
        exportExcel.setDataList(list);
        exportExcel.writeFile("员工信息", req, resp);*/

        // 导出分组为 1 的 Excel 模板
        /*ExportExcel exportExcel2 = new ExportExcel("员工信息模板", User.class, 2, 1);
        exportExcel2.writeFile("员工信息", req, resp);*/

        // 导出用户信息
        /*ExportExcel exportExcel = new ExportExcel("员工信息模板", User.class, 1);
        exportExcel.setDataList(list);
        exportExcel.writeFile("员工信息", req, resp);*/

        // 导出分组为 1 的用户信息
        /*ExportExcel exportExcel = new ExportExcel("员工信息模板", User.class, 1, 1);
        exportExcel.setDataList(list);
        exportExcel.writeFile("员工信息", req, resp);*/

        // 导出手动添加的用户信息
        /*ExportExcel exportExcel = new ExportExcel("员工信息模板", User.class, 1);
        Row row = exportExcel.addRow();
        exportExcel.addCell(row, 0, "test", 2, String.class);
        exportExcel.addCell(row, 1, 33, 2, Integer.class);
        exportExcel.addCell(row, 2, new Date(), 2, Date.class);
        exportExcel.addCell(row, 3, "上海", 2, String.class);
        exportExcel.writeFile("员工信息", req, resp);*/

    }
}
