package com.mobin.CSVToExcel.finalexcel;

import jxl.write.WriteException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
public class 交叉 {
    public static void main(String[] args) throws IOException, WriteException {

//        FileInputStream file = new FileInputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\工单ww.xlsx"));
//        ZipSecureFile.setMinInflateRatio(-1.0d);
//
//        XSSFWorkbook workbook = new  XSSFWorkbook(file);
//        XSSFSheet sheet = workbook.getSheet("ce");

        FileInputStream file1 = new FileInputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\交叉-w33.xlsx"));
        XSSFWorkbook workbook1 = new  XSSFWorkbook(file1);
        XSSFSheet sheet1 = workbook1.getSheet("Sheet1");


        //总体的服务率贡献wow
        Cell celltemp1 = sheet1.getRow(5).getCell(50);
        Cell celltemp2 = sheet1.getRow(5).getCell(51);
        Cell celltemp = sheet1.getRow(5).getCell(53);

        int col=celltemp.getColumnIndex()-2;
        int col2=celltemp.getColumnIndex()-3;

        int row=celltemp.getRowIndex();
        Cell temp1=sheet1.getRow(row).getCell(col);
        Cell temp2=sheet1.getRow(row).getCell(col2);
        int num =  sheet1.getRow(2).getLastCellNum();
        System.out.println(celltemp.getAddress());
        celltemp.setCellFormula(temp1.getAddress().toString()+"-"+temp2.getAddress().toString());

        //处理服务率贡献wow-33该列
        Cell celltempmeifaWow = sheet1.getRow(18).getCell(35);
        int col3=celltempmeifaWow.getColumnIndex()-1;
        int col4=celltempmeifaWow.getColumnIndex()-2;
        int row3=celltempmeifaWow.getRowIndex();
        Cell temp3=sheet1.getRow(row3).getCell(col3);
        Cell temp4=sheet1.getRow(row3).getCell(col4);
        celltempmeifaWow.setCellFormula(temp3.getAddress().toString()+"-"+temp4.getAddress().toString());

        //处理服务率贡献wow-33该列
        Cell celltempmeifaWow1 = sheet1.getRow(30).getCell(35);
        int col5=celltempmeifaWow1.getColumnIndex()-1;
        int col6=celltempmeifaWow1.getColumnIndex()-2;
        int row4=celltempmeifaWow1.getRowIndex();
        Cell temp5=sheet1.getRow(row4).getCell(col5);
        Cell temp6=sheet1.getRow(row4).getCell(col6);
        celltempmeifaWow1.setCellFormula(temp5.getAddress().toString()+"-"+temp6.getAddress().toString());
        //处理服务率贡献wow-33该列
        Cell celltempmeifaWow2 = sheet1.getRow(41).getCell(35);
        int col7=celltempmeifaWow2.getColumnIndex()-1;
        int col8=celltempmeifaWow2.getColumnIndex()-2;
        int row5=celltempmeifaWow2.getRowIndex();
        Cell temp7=sheet1.getRow(row5).getCell(col7);
        Cell temp8=sheet1.getRow(row5).getCell(col8);
        celltempmeifaWow2.setCellFormula(temp7.getAddress().toString()+"-"+temp8.getAddress().toString());
        //处理服务率贡献wow-33该列
        Cell celltempmeifaWow3 = sheet1.getRow(52).getCell(35);
        int col9=celltempmeifaWow3.getColumnIndex()-1;
        int col10=celltempmeifaWow3.getColumnIndex()-2;
        int row6=celltempmeifaWow3.getRowIndex();
        Cell temp9=sheet1.getRow(row6).getCell(col9);
        Cell temp10=sheet1.getRow(row6).getCell(col10);
        celltempmeifaWow3.setCellFormula(temp9.getAddress().toString()+"-"+temp10.getAddress().toString());




        //指定生成文件位置
        FileOutputStream output_file = new FileOutputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\"+"不满意ww.xlsx"));
        workbook1.write(output_file);//write excel document to output stream
        output_file.close(); //close the file
    }


}
