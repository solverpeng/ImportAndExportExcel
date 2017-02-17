package com.solverpeng.poi.beans;

import com.solverpeng.poi.utils.excel.ExcelField;

import java.util.Date;

/**
 * Created by solverpeng on 2017/2/16 0016.
 */
public class User {
    private String id;
    @ExcelField(title = "用户姓名**User Name", groups = {1, 2}, sort = 1, align = 2)
    private String userName;
    @ExcelField(title = "用户年龄**User Age", groups = {1}, sort = 2, align = 2)
    private Integer age;
    @ExcelField(title = "用户生日**User Birth", groups = {2}, sort = 3, align = 2)
    private Date birth;
    @ExcelField(title = "居住地址**Address", groups = {1, 2}, sort = 4, align = 2)
    private String address;

    public User() {
    }

    public User(String id, String userName, Integer age, Date birth, String address) {
        this.id = id;
        this.userName = userName;
        this.age = age;
        this.birth = birth;
        this.address = address;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Date getBirth() {
        return birth;
    }

    public void setBirth(Date birth) {
        this.birth = birth;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        User user = (User) o;

        if (id != null ? !id.equals(user.id) : user.id != null) return false;
        if (userName != null ? !userName.equals(user.userName) : user.userName != null) return false;
        if (age != null ? !age.equals(user.age) : user.age != null) return false;
        if (birth != null ? !birth.equals(user.birth) : user.birth != null) return false;
        return address != null ? address.equals(user.address) : user.address == null;
    }

    @Override
    public int hashCode() {
        int result = id != null ? id.hashCode() : 0;
        result = 31 * result + (userName != null ? userName.hashCode() : 0);
        result = 31 * result + (age != null ? age.hashCode() : 0);
        result = 31 * result + (birth != null ? birth.hashCode() : 0);
        result = 31 * result + (address != null ? address.hashCode() : 0);
        return result;
    }

    @Override
    public String toString() {
        return "User{" +
                "id='" + id + '\'' +
                ", userName='" + userName + '\'' +
                ", age=" + age +
                ", birth=" + birth +
                ", address='" + address + '\'' +
                '}';
    }
}
