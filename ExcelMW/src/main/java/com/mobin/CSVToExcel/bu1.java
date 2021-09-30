package com.mobin.CSVToExcel;



//import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class bu1 {

    public static void main(String[] args) throws IOException {
        //创建一个模板文件，里面自己手动填入数据
        FileInputStream file = new FileInputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\bu1.xlsx"));
        XSSFWorkbook workbook = new  XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //获取sheet页数据行数
        int num =  sheet.getLastRowNum()+1;

        XSSFSheet sheet1 = workbook.createSheet("ce");
        //手动填充数据的区域
        //AreaReference a=new AreaReference("A1:E4");
        AreaReference a=new AreaReference("A1:I"+num+"",SpreadsheetVersion.EXCEL2007);
        //AreaReference a=new AreaReference("A1:E4",SpreadsheetVersion.EXCEL2007);
        //数据透视表生成为位置
        CellReference b=new CellReference("I5");
        //生成数据透视图
        XSSFPivotTable pivotTable = sheet1.createPivotTable(a,b,sheet);
        //添加行标签
//        pivotTable.addRowLabel(6);
//        pivotTable.addRowLabel(7);
        //添加筛选项
        pivotTable.addReportFilter(1);
        pivotTable.addReportFilter(2);
        System.out.println(pivotTable.getCTPivotTableDefinition().getFilters());
        CellAddress address = new CellAddress("L8");
//        ArrayList<String> a1=new ArrayList<>();
//        String[] s1={"丽人","美发","美容美体","医疗整形","美甲","瑜伽","医疗","口腔齿科","体检中心","null"};
//        for(int i=0;i<s1.length;i++){
//            sheet1.createRow(6+i).createCell(11);
//            Cell cell = sheet1.getRow(6+i).getCell(11);
//            cell.setCellValue(s1[i]);
//        }
//        sheet1.getRow(6).createCell(12);
//        Cell cell1 = sheet1.getRow(6).getCell(12);
//        cell1.setCellFormula("VLOOKUP(L7,$I$6:$J$35,2,FALSE)");

//        String[] s2={"C端门店展示问题","POI信息及收录问题","产品功能设置及咨询","诚信问题","广告费用退款问题","合作问题咨询","活动问题","技师问题","结算问题","其他","投诉销售","项目上下线及变更","销售服务[纠纷/不满意]","协助寻找会员","续约问题","账号问题","断线","合计"};
//        for(int i=0;i<s2.length;i++){
//            sheet1.createRow(26+i).createCell(11);
//            Cell cell = sheet1.getRow(26+i).getCell(11);
//            cell.setCellValue(s2[i]);
//        }
//        sheet1.getRow(7).createCell(12);
//        Cell cell1 = sheet1.getRow(7).getCell(12);
//        cell1.setCellFormula("VLOOKUP(L8,$I$8:$K$25,2,FALSE)");
//        sheet1.getRow(7).createCell(13);
//        Cell cell2 = sheet1.getRow(7).getCell(13);
//        cell2.setCellFormula("VLOOKUP(L8,$I$8:$K$25,3,FALSE)");
//
//        sheet1.getRow(26).createCell(12);
//        Cell cell3 = sheet1.getRow(26).getCell(12);
//        cell3.setCellFormula("VLOOKUP(L27,$I$27:$K$43,2,FALSE)");
//        sheet1.getRow(26).createCell(13);
//        Cell cell4 = sheet1.getRow(26).getCell(13);
//        cell4.setCellFormula("VLOOKUP(L27,$I$27:$K$43,3,FALSE)");

//        XSSFRow row = sheet1.getRow(address.getRow());//得到行
//        XSSFCell cell = row.getCell(address.getColumn());//得到列
//        cell.setCellValue("213");
        //workbook.getSheet("ce").getCellComment(CellAddress).
//                CellReference c1=new CellReference("I5");
//        XSSFCell newNameCell = c1.createCell(cellIndex++, Cell.CELL_TYPE_STRING);
        //添加列数据
        //pivotTable.addDataColumn(0, true);
//        pivotTable.addDataColumn(1, true);
//        pivotTable.addDataColumn(2,false);
        pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE,3,"平均值项:万订单服务率");
        pivotTable.addColumnLabel(DataConsolidateFunction.SUM,4,"求和项:千POI服务率");
        pivotTable.addColumnLabel(DataConsolidateFunction.SUM,5,"求和项:履约发生率");
        pivotTable.addColumnLabel(DataConsolidateFunction.SUM,7,"求和项:不满意度（1分+2分）");

//        for (int i = 0; i < 4; i++) {
//            CTPivotField ctPivotField = pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(i);
//            ctPivotField.setOutline(false);
//        }
        //指定生成文件位置
        FileOutputStream output_file = new FileOutputStream(new File("D:\\codemt\\CSVToExcel-master\\src\\main\\java\\com\\mobin\\CSVToExcel\\"+"bu1ww.xlsx"));
        workbook.write(output_file);//write excel document to output stream
        output_file.close(); //close the file
    }
}



