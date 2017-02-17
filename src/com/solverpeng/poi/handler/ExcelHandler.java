package com.solverpeng.poi.handler;

import com.solverpeng.poi.beans.User;
import com.solverpeng.poi.utils.excel.ImportExcel;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

/**
 * Created by solverpeng on 2017/2/17 0017.
 */
@Controller
public class ExcelHandler {
    @RequestMapping(value = "/importExcel", method = RequestMethod.POST)
    public String importExcel(@RequestParam("file") MultipartFile file) throws IOException, InstantiationException, IllegalAccessException {
        ImportExcel importExcel = new ImportExcel(file, 1, 0);
        List<User> userList = importExcel.getDataList(User.class);
        for (User user : userList) {
            System.out.println(user);
        }
        return "success";
    }
}
