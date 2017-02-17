package com.solverpeng.poi.read;

import com.solverpeng.poi.beans.User;
import com.solverpeng.poi.utils.excel.ImportExcel;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * Created by solverpeng on 2017/2/17 0017.
 */
public class ImportExcelTest {
    public static void main(String[] args) throws IOException, InstantiationException, IllegalAccessException {
        String fileName = "out/artifacts/poi_03_war_exploded/files/员工信息.xls";
        String fileName2 = "out/artifacts/poi_03_war_exploded/files/员工信息.xlsx";
        File file = new File(fileName);
        File file2 = new File(fileName2);
        ImportExcel importExcel = new ImportExcel(file, 1);
        ImportExcel importExcel2 = new ImportExcel(file2, 1);
        List<User> userList = importExcel.getDataList(User.class);
        for (User user : userList) {
            System.out.println(user);
        }
    }
}
