package com.mobin.CSVToExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class asea {
    public static void main(String[] args) throws ParseException, IOException {
        SimpleDateFormat simpleFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        Date fromDate2 = simpleFormat.parse("2021-08-27 13:49:25");
        Date toDate2 = simpleFormat.parse("2021-05-12 01:54:20");
        long from2 = fromDate2.getTime();
        long to2 = toDate2.getTime();
        int hours = (int) ((to2 - from2) / (1000 * 60 * 60));
        System.out.println("两个时间之间的小时差为：" + hours);


        FileInputStream file = new FileInputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\时差.xlsx"));
        XSSFWorkbook workbook = new  XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(2);
        int lastRowNum = sheet.getLastRowNum();

        Cell cell1 = sheet.getRow(1).getCell(5);
        Cell cell2 = sheet.getRow(1).getCell(6);
        Cell cell3 = sheet.getRow(1).createCell(8);
        cell3.setCellValue(hours);

        FileOutputStream output_file = new FileOutputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\"+"时差ww.xlsx"));
        workbook.write(output_file);//write excel document to output stream
        output_file.close(); //close the file



    }
}
