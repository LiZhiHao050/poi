package com.example.write;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class WriteExcel {

    public static void main(String[] args) throws Exception {
        // 1.创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 2.创建工作表
        XSSFSheet sheet = workbook.createSheet("SHEET1");
        // 3.创建行
        XSSFRow row = sheet.createRow(0);

        // 4.创建单元格
        row.createCell(0).setCellValue("123");
        row.createCell(1).setCellValue("456");

        // 输出流输出文件
        FileOutputStream fos = new FileOutputStream("Z:/Pic/output.xlsx");
        workbook.write(fos);

        fos.flush();

        // 关闭资源(先开后关)
        fos.close();
        workbook.close();

        System.out.println("文件已输出");
    }

}
