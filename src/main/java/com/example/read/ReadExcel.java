package com.example.read;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ReadExcel {

    public static void main(String[] args) throws IOException {
        // 1.获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook("Z:/Pic/test.xlsx");
        // 2.获取sheet表
        XSSFSheet sheet = workbook.getSheet("sheet1");// 根据名称获取
//        XSSFSheet sheet = workbook.getSheetAt(0);     // 根据位置获取,0是第一个
        // 3.获取行
        for (Row row : sheet) {
//            System.out.println("Row:" + row);
            // 4.获取单元格
            for (Cell cell : row) {
//                System.out.println("Cell:" + cell);
                // 5.获取单元格中的内容
                String value = cell.getStringCellValue();
                System.out.println(value);
            }
        }

        // 普通for循环
        // 获取最后一行的索引
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null) {
                // 获取单元格最后索引
                short lastCellNum = row.getLastCellNum();
                for (int j = 0; j <= lastCellNum ; j++) {
                    XSSFCell cell = row.getCell(j);
                    if (cell != null) {
                        String value = cell.getStringCellValue();
                        System.out.println(value);
                    }
                }
            }
        }

        // 释放资源
        workbook.close();
    }

}
