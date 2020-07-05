package com.example.read;

import com.example.entity.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ReadListExcel {

    public static void main(String[] args) throws IOException {
        // 1.获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook("Z:/Pic/User.xlsx");
        // 2.获取工作表
        XSSFSheet sheet = workbook.getSheet("用户表");
        // 定义对象存储集合
        List<User> userList = new ArrayList<>();
        // 3.获取表的所有行
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            // 对行进行判空
            if (row != null) {
                // 定义单元格值存储集合
                List<String> cellValue = new ArrayList<>();
                // 遍历单元格
                for (Cell cell : row) {
                    // 对单元格进行判空
                    if (cell != null) {
                        // 设置接收单元格内容的类型,防止出现类型转换异常
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        String value = cell.getStringCellValue();
                        if (value != null && !"".equals(value)) {   // 判空操作
                            cellValue.add(value);
                        }
                    }
                }

                // 判空操作
                if (cellValue.size() > 0) {
                    // 创建对象设置相应值并存入集合
                    User user = new User(Integer.parseInt(cellValue.get(0)), cellValue.get(1),
                            cellValue.get(2), Integer.parseInt(cellValue.get(3)), cellValue.get(4));
                    userList.add(user);
                }

            }
        }

        // 获取集合所有值
        for (User user : userList) {
            System.out.println(user);
        }

    }

}
