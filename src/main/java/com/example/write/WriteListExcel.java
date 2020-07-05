package com.example.write;

import com.example.entity.User;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class WriteListExcel {

    public static void main(String[] args) throws Exception {
        /*List<User> userList = new ArrayList<>();
        for (int i = 1; i <= 10; i++) {
            userList.add(new User(i, "name" + i, "男", 20 + i, "000" + i));
        }

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("用户表");

        XSSFRow row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("ID");
        row0.createCell(1).setCellValue("姓名");
        row0.createCell(2).setCellValue("性别");
        row0.createCell(3).setCellValue("年龄");
        row0.createCell(4).setCellValue("身份证号");

        for (int i = 0; i < userList.size(); i++) {
            User user = userList.get(i);
            XSSFRow row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getName());
            row.createCell(2).setCellValue(user.getGender());
            row.createCell(3).setCellValue(user.getAge());
            row.createCell(4).setCellValue(user.getIdCard());
        }

        FileOutputStream fos = new FileOutputStream("Z:/Pic/User.xlsx");
        workbook.write(fos);

        fos.flush();
        fos.close();

        workbook.close();
        System.out.println("写入完成!");*/
        styleExcel();

    }


    public static void styleExcel() throws Exception {
        List<User> userList = new ArrayList<>();
        for (int i = 1; i <= 10; i++) {
            userList.add(new User(i, "name" + i, "男", 20 + i, "00000000000000000000" + i));
        }

        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建单元格样式对象
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());        // 设置颜色
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);             // 设置规则,实心

        // 创建字体样式
        XSSFFont font = workbook.createFont();
        font.setFontName("微软雅黑");                     // 设置字体
        font.setColor(IndexedColors.PINK.getIndex());   // 设置颜色
        font.setFontHeight(12);                         // 设置字体大小
        font.setItalic(false);                           // 是否使用斜体
        font.setStrikeout(false);                        // 是否使用划线

        // 将字体样式放入单元格样式中
        cellStyle.setFont(font);


        XSSFSheet sheet = workbook.createSheet("用户表");

        XSSFRow row0 = sheet.createRow(0);

        XSSFCell cell0 = row0.createCell(0);
        cell0.setCellValue("ID");
        cell0.setCellStyle(cellStyle);


        XSSFCell cell1 = row0.createCell(1);
        cell1.setCellValue("姓名");
        cell1.setCellStyle(cellStyle);

        XSSFCell cell2 = row0.createCell(2);
        cell2.setCellValue("性别");
        cell2.setCellStyle(cellStyle);

        XSSFCell cell3 = row0.createCell(3);
        cell3.setCellValue("年龄");
        cell3.setCellStyle(cellStyle);

        XSSFCell cell4 = row0.createCell(4);
        cell4.setCellValue("身份证号");
        cell4.setCellStyle(cellStyle);


        for (int i = 0; i < userList.size(); i++) {
            User user = userList.get(i);
            XSSFRow row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getName());
            row.createCell(2).setCellValue(user.getGender());
            row.createCell(3).setCellValue(user.getAge());
            row.createCell(4).setCellValue(user.getIdCard());
            sheet.autoSizeColumn(i);        // 自动调整列宽(第几列)
        }

        FileOutputStream fos = new FileOutputStream("Z:/Pic/User.xlsx");
        workbook.write(fos);

        fos.flush();
        fos.close();

        workbook.close();
        System.out.println("写入完成!");
    }

}
