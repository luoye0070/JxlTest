package com.lj.jxl;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created with IntelliJ IDEA.
 * User: Administrator
 * Date: 15-5-20
 * Time: 下午9:19
 * To change this template use File | Settings | File Templates.
 */
public class ExcelImport {
    public String[] workbookFullNames=null;
    public String excelFilePath=null;
    public String charsetName=null;
    public String dateFormtStr=null;
    public ExcelImport(String excelFilePath,String charsetName,String dateFormtStr){
        this.excelFilePath=excelFilePath;
        this.charsetName=charsetName;
        this.dateFormtStr=dateFormtStr;
    }
    public ExcelImport(String excelFilePath,String charsetName,String dateFormtStr,String[] workbookFullNames){
        this.excelFilePath=excelFilePath;
        this.charsetName=charsetName;
        this.dateFormtStr=dateFormtStr;
        this.workbookFullNames=workbookFullNames;
    }
    public ExcelImport(String excelFilePath){
        this.excelFilePath=excelFilePath;
    }
    public ExcelImport(String excelFilePath,String[] workbookFullNames){
        this.excelFilePath=excelFilePath;
        this.workbookFullNames=workbookFullNames;
    }
    public ExcelImport(){
    }

    public int importData(String txtFilePath,String sheetName,int beginRow,int beginCell,String decollator){
        int result=0;
        File excelFile=new File(excelFilePath);
        if(!excelFile.exists()){
            return -1;//excel文件不存在
        }
        File txtFile=new File(txtFilePath);
        if(!txtFile.exists()){
            return -2;//txt文件不存在
        }
        if(charsetName==null||charsetName.equals("")){
            charsetName="gb2312";
        }
        if(decollator==null||decollator.equals("")){
            decollator="\\|";
        }
        BufferedReader bufferedReader=null;
        Workbook workbook=null;
        try{
            workbook=new HSSFWorkbook(new FileInputStream(excelFilePath));
            Sheet sheet= workbook.getSheet(sheetName);
            if(sheet==null){
                return -3;
            }
            //读取数据
            bufferedReader=new BufferedReader(new InputStreamReader(new FileInputStream(txtFile),charsetName));
            String lineStr=null;
            int rowIndex=beginRow;
            while ((lineStr=bufferedReader.readLine())!=null){
                System.out.println("lineStr->"+lineStr);
                String[] cellStrs=lineStr.split(decollator);
                System.out.println("cellStrs-->"+":");
                //将数据写入excel
                Row row=sheet.getRow(rowIndex);
                if(row==null){
                    row=sheet.createRow(rowIndex);
                }
                for(int i=0;i<cellStrs.length;i++){
                    System.out.print("-"+cellStrs[i]);
                    int cellIndex=beginCell+i;
                    Cell cell=row.getCell(cellIndex);
                    if(cell==null){
                        cell=row.createCell(cellIndex);
                    }
                    SimpleDateFormat dateFormat=new SimpleDateFormat(this.dateFormtStr);
                    try {
                        Date date=dateFormat.parse(cellStrs[i]);
                        System.out.println("date="+date);
                        cell.setCellValue(date);
                        CellStyle cellStyle=workbook.createCellStyle();
                        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy/M/d"));
                        cell.setCellStyle(cellStyle);
                    }catch (ParseException ex){
                        try {
                            double numericalValue=Double.parseDouble(cellStrs[i]);
                            cell.setCellValue(numericalValue);
                            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        }catch (NumberFormatException ex1){
                            cell.setCellValue(cellStrs[i]);
                            cell.setCellType(HSSFCell.ENCODING_UNCHANGED);
                        }
                    }
//                    cell.setCellValue(cellStrs[i]);
//                    cell.setCellType(HSSFCell.ENCODING_UNCHANGED);
                }
                System.out.println("");
                rowIndex++;
            }
            //sheet.setForceFormulaRecalculation(true);
            //HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.write(new FileOutputStream(excelFile));
        }catch (Exception ex){
            ex.printStackTrace();
            return -100;
        }finally {
            try {
                if(bufferedReader!=null){
                    bufferedReader.close();
                }
                if(workbook!=null){
                    workbook.close();
                }
            }catch (Exception ex){
               ex.printStackTrace();
            }
        }

        return result;
    }

    public int cleanNullValue(String sheetName,int beginRow,int beginCell,int endRow,int endCell){
        int result=0;
        File excelFile=new File(excelFilePath);
        if(!excelFile.exists()){
            return -1;//excel文件不存在
        }
        if(charsetName==null||charsetName.equals("")){
            charsetName="gb2312";
        }
        HSSFWorkbook workbook=null;
        try{
            workbook=new HSSFWorkbook(new FileInputStream(excelFilePath));
            HSSFFormulaEvaluator evaluator=new HSSFFormulaEvaluator(workbook);
            Sheet sheet= workbook.getSheet(sheetName);
            if(sheet==null){
                return -3;
            }
            int rowIndex=beginRow;
            while (rowIndex<=endRow){
                //将数据写入excel
                Row row=sheet.getRow(rowIndex);
                if(row==null){
                    continue;
                }
                for(int i=beginCell;i<=endCell;i++){
                    int cellIndex=beginCell;
                    Cell cell=row.getCell(cellIndex);
                    if(cell==null){
                        continue;
                    }

                    System.out.println(rowIndex+"*"+cellIndex);
                    if(cell.getCellType()==Cell.CELL_TYPE_FORMULA){
                        System.out.println("is Foumula");
                        evaluator.evaluateInCell(cell);
                        try{
                            double value=cell.getNumericCellValue();
                            System.out.println("value is "+value);
                        }catch (IllegalStateException ex){
                            cell=row.createCell(cellIndex);
                            cell.setCellValue(0);
                            cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                            System.out.println("set 0");
                        }
                    }else{
                        System.out.println("is not Formula");
                    }
//                    cell.setCellValue(cellStrs[i]);
//                    cell.setCellType(HSSFCell.ENCODING_UNCHANGED);
                }
                rowIndex++;
            }
            //sheet.setForceFormulaRecalculation(true);
            //HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.write(new FileOutputStream(excelFile));
        }catch (Exception ex){
            ex.printStackTrace();
            return -100;
        }finally {
            try {
                if(workbook!=null){
                    workbook.close();
                }
            }catch (Exception ex){
                ex.printStackTrace();
            }
        }

        return result;
    }
    public int evaluateAllFormulaCells(){
        int result=0;
        File excelFile=new File(excelFilePath);
        if(!excelFile.exists()){
            return -1;//excel文件不存在
        }
        HSSFWorkbook workbook=null;
        try{
            workbook=new HSSFWorkbook(new FileInputStream(excelFilePath));
            HSSFFormulaEvaluator hssfFormulaEvaluator=new HSSFFormulaEvaluator(workbook);
            if(workbookFullNames!=null&&workbookFullNames.length>0){
                int length=workbookFullNames.length;
                String [] workbookNames=new String[length+1];
                workbookNames[0]=getFileName(excelFilePath);
                HSSFFormulaEvaluator[] hssfFormulaEvaluators=new HSSFFormulaEvaluator[length+1];
                hssfFormulaEvaluators[0]=hssfFormulaEvaluator;
                for(int i=0;i<length;i++){
                    workbookNames[i+1]=getFileName(workbookFullNames[i]);
                    hssfFormulaEvaluators[i+1]=new HSSFFormulaEvaluator(new HSSFWorkbook(new FileInputStream(workbookFullNames[i])));
                }
               HSSFFormulaEvaluator.setupEnvironment(workbookNames,hssfFormulaEvaluators);
            }
            hssfFormulaEvaluator.setIgnoreMissingWorkbooks(true);
            //hssfFormulaEvaluator.clearAllCachedResultValues();
            hssfFormulaEvaluator.evaluateAll();
            //HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.write(new FileOutputStream(excelFile));
        }catch (Exception ex){
            ex.printStackTrace();
            return -100;
        }finally {
            try {
                if(workbook!=null){
                    workbook.close();
                }
            }catch (Exception ex){
                ex.printStackTrace();
            }
        }

        return result;
    }
    private String getFileName(String fileFullName){
        if(fileFullName==null){
            return null;
        }
        int idx=fileFullName.lastIndexOf("/");
        if(idx==-1){
            return fileFullName;
        }
        String fileName=fileFullName.substring(idx+1,fileFullName.length());
        System.out.println("fileName->"+fileName);
        return fileName;
    }
}
