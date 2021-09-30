package com.mobin.CSVToExcel.monthPack;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class 处理月工单 {
    public static void main(String[] args) throws IOException {
        //创建一个模板文件，里面自己手动填入数据
        FileInputStream file = new FileInputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\月工单预处理ww.xlsx"));
        XSSFWorkbook workbook = new  XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        //获取sheet页数据行数
        int num =  sheet.getLastRowNum()+1;
        System.out.println(num);
        XSSFSheet sheet1 = workbook.createSheet("表1");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable = sheet1.createPivotTable(a,b,sheet);
//        //添加行标签
        pivotTable.addRowLabel(58);

        pivotTable.addReportFilter(59);
        pivotTable.addReportFilter(61);
//        pivotTable.addReportFilter(7);
        pivotTable.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");


        //表2
        XSSFSheet sheet2 = workbook.createSheet("表2");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a2=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b2=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable2 = sheet2.createPivotTable(a2,b2,sheet);
//        //添加行标签
        pivotTable2.addRowLabel(60);

//        pivotTable.addReportFilter(63);
//        pivotTable.addReportFilter(65);
//        pivotTable.addReportFilter(27);
        pivotTable2.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");

        //图1及表4
        XSSFSheet sheet3 = workbook.createSheet("图1及表4");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a3=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b3=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable3 = sheet3.createPivotTable(a3,b3,sheet);
//        //添加行标签
        pivotTable3.addRowLabel(58);

//        pivotTable.addReportFilter(63);
//        pivotTable.addReportFilter(65);
//        pivotTable.addReportFilter(27);
        pivotTable3.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");

        //附件1以及表3
        XSSFSheet sheet4 = workbook.createSheet("附件1及表3");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a4=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b4=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable4 = sheet4.createPivotTable(a4,b4,sheet);
//        //添加行标签
        pivotTable4.addRowLabel(58);
        pivotTable4.addRowLabel(10);
//        pivotTable.addReportFilter(63);
        pivotTable4.addReportFilter(61);
//        pivotTable.addReportFilter(27);
        pivotTable4.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");


        //规则外退款
        XSSFSheet sheet5 = workbook.createSheet("规则外退款");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a5=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b5=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable5 = sheet5.createPivotTable(a5,b5,sheet);
//        //添加行标签
        pivotTable5.addRowLabel(59);
        pivotTable5.addRowLabel(60);
//        pivotTable.addReportFilter(63);
        pivotTable5.addReportFilter(58);
        pivotTable5.addColLabel(61);
//        pivotTable.addReportFilter(27);
        pivotTable5.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");

        //附件3
        XSSFSheet sheet6 = workbook.createSheet("附件3");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a6=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b6=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable6 = sheet6.createPivotTable(a6,b6,sheet);
//        //添加行标签
        pivotTable6.addRowLabel(60);
        pivotTable6.addRowLabel(58);
//        pivotTable.addReportFilter(63);
        //pivotTable6.addReportFilter(62);
        //pivotTable6.addColLabel(65);
//        pivotTable.addReportFilter(27);
        pivotTable6.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");

        //附件4
        XSSFSheet sheet7 = workbook.createSheet("附件4");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a7=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b7=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable7 = sheet7.createPivotTable(a7,b7,sheet);
//        //添加行标签
        pivotTable7.addRowLabel(58);
        pivotTable7.addRowLabel(10);
        pivotTable7.addRowLabel(15);
//        pivotTable.addReportFilter(63);
        pivotTable7.addReportFilter(61);
        //pivotTable6.addColLabel(65);
//        pivotTable.addReportFilter(27);
        pivotTable7.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");


        //附件5
        XSSFSheet sheet8 = workbook.createSheet("附件5");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a8=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b8=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable8 = sheet8.createPivotTable(a8,b8,sheet);
//        //添加行标签
        pivotTable8.addRowLabel(18);
//        pivotTable8.addRowLabel(14);
//        pivotTable8.addRowLabel(19);
//        pivotTable.addReportFilter(63);
        pivotTable8.addReportFilter(61);
        pivotTable8.addReportFilter(58);
        //pivotTable6.addColLabel(65);
//        pivotTable.addReportFilter(27);
        pivotTable8.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");

//表5
        XSSFSheet sheet9 = workbook.createSheet("表5");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a9=new AreaReference("A1:BJ"+num+"", SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b9=new CellReference("B5");
        //生成数据透视图
        XSSFPivotTable pivotTable9 = sheet9.createPivotTable(a9,b9,sheet);
//        //添加行标签
        pivotTable9.addRowLabel(10);
//        pivotTable8.addRowLabel(14);
//        pivotTable8.addRowLabel(19);
//        pivotTable.addReportFilter(63);
        pivotTable9.addReportFilter(58);
        pivotTable9.addReportFilter(61);
        //pivotTable6.addColLabel(65);
//        pivotTable.addReportFilter(27);
        pivotTable9.addColumnLabel(DataConsolidateFunction.COUNT,0,"计数项:工单ID");

















        FileOutputStream output_file = new FileOutputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\"+"月工单处理ww.xlsx"));
        workbook.write(output_file);//write excel document to output stream
        output_file.close(); //close the file
    }
}
