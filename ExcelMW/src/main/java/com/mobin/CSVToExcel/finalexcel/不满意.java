package com.mobin.CSVToExcel.finalexcel;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class 不满意 {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\场景不满意ww.xlsx"));
        ZipSecureFile.setMinInflateRatio(-1.0d);

        XSSFWorkbook workbook = new  XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("ce");

        FileInputStream file1 = new FileInputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\不满意-w33.xlsx"));
        XSSFWorkbook workbook1 = new  XSSFWorkbook(file1);
        XSSFSheet sheet1 = workbook1.getSheet("客服");

        int num =  sheet1.getRow(2).getLastCellNum();
        for(int i=0;i<1;i++){
            //sheet.getRow(7).getCell(12);
            Cell celltemp = sheet.getRow(8+i).getCell(12);
            XSSFFormulaEvaluator eva= new XSSFFormulaEvaluator(workbook);

            FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(celltemp);
            CellValue cellValue = evaluator.evaluate(celltemp);
            Double celldata = cellValue.getNumberValue();
            System.out.println(celltemp.getCellFormula());
            sheet1.getRow(2+i).createCell(52);
            Cell cell1 = sheet1.getRow(2+i).getCell(52);
            //cell1.;
            copyCell(sheet,8,12,sheet1,2,53);
        }

        //指定生成文件位置
        FileOutputStream output_file = new FileOutputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\"+"不满意ww.xlsx"));
        workbook1.write(output_file);//write excel document to output stream
        output_file.close(); //close the file
    }

    /**从(s1,r1,c1)单元格复制内容到(s2,r2,c2)*/
    public static void copyCell(Sheet s2, int r2, int c2, Sheet s1, int r1, int c1) {

        Cell cell = getCell(s2, r2, c2);

        Object obj= getObj(s1, r1, c1);
        if(null == obj){
            //为空不处理
        }else if (obj instanceof String) {
            cell.setCellValue((String) obj);
        } else if (obj instanceof Date) {
            cell.setCellValue((Date) obj);
        } else if (obj instanceof Double) {
            cell.setCellValue((double) obj);
        } else {
            System.out.println("未处理类型：" + obj.getClass());
        }
    }

    private static Cell getCell(Sheet sheet, int r, int c) {

        Row row = sheet.getRow(r);
        if (row == null) {
            row = sheet.createRow(r);
        }

        Cell cell = row.getCell(c);
        if (cell == null) {
            cell = row.createCell(c);
        }

        return cell;
    }

    private static Object getObj(Sheet sheet, int r, int c) {

        Cell cell = getCell(sheet, r, c);

        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            } else {
                return cell.getNumericCellValue();
            }
        } else if (cell.getCellTypeEnum() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellTypeEnum() == CellType.BLANK) {
            return null;
        } else if (cell.getCellTypeEnum() == CellType.FORMULA) {
            if(cell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC){
                return cell.getNumericCellValue();
            }else{
                return cell.getStringCellValue();
            }
        } else {
            System.out.println("("+r+","+c+")未处理类型：" + cell.getCellTypeEnum());
        }

        return "";
    }
}
