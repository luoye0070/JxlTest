package com.lj.jxl;

//import jxl.Workbook;
//import jxl.write.Label;
//import jxl.write.WritableSheet;
//import jxl.write.WritableWorkbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.channels.FileChannel;

/**
 * Created with IntelliJ IDEA.
 * User: Administrator
 * Date: 15-5-18
 * Time: 下午9:26
 * To change this template use File | Settings | File Templates.
 */
public class JxlTest {
    public static void main(String args[]){
        System.out.println("ddddd");
        try{
            System.out.println(new File("").getAbsolutePath());
//            Workbook workbook=Workbook.getWorkbook(new File("日报基础2015..xls"));
//            WritableWorkbook writableWorkbook=Workbook.createWorkbook(new File("日报基础2015_d..xls"),workbook);
//         WritableSheet sheet= writableWorkbook.getSheet("当日赔款");
//            sheet.addCell(new Label(0,0,"ddd"));
//            writableWorkbook.write();
//            writableWorkbook.close();


            File file=new File("日报基础2015..xls");
            File file1=new File("日报基础2015_d..xls");
            FileOutputStream fileOutputStream=new FileOutputStream(file1);
            FileChannel foutChannel=fileOutputStream.getChannel();
            FileInputStream fileInputStream=new FileInputStream(file);
            FileChannel finChannel=fileInputStream.getChannel();

            finChannel.transferTo(0,finChannel.size(),foutChannel);
            fileInputStream.close();
            finChannel.close();
            fileOutputStream.close();;
            foutChannel.close();

            System.out.println(new ExcelImport("日报基础2015_d..xls",null,"yyyy-MM-dd").importData("当年当日签单保费.txt","当日签单",1,0,null));

//            Workbook workbook=new HSSFWorkbook(new FileInputStream("日报基础2015..xls"));
//            Sheet sheet= workbook.getSheet("当日签单");
//            Row row=sheet.getRow(95);
//            Cell cellG=row.getCell(11);
//            System.out.println("cell:"+cellG.getNumericCellValue());
//
//
//            Cell cell =row.createCell(11);
//            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
//
//            cell.setCellValue(12090L);
//            CellStyle cellStyle=workbook.createCellStyle();
//            cellStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
//            cell.setCellStyle(cellStyle);
//
//            workbook.write(new FileOutputStream("ddd.xls"));
//            workbook.close();
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }
}