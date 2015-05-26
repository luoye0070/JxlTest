package com.lj.jxl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.channels.FileChannel;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created with IntelliJ IDEA.
 * User: Administrator
 * Date: 15-5-25
 * Time: 下午8:24
 * To change this template use File | Settings | File Templates.
 */
public class FileScanAndImport {
    public final static String templateFileDetail="日报基础2015..xls";
    public final static String templateFileGeneral="业务动态.xls";
    public final static HashMap<String,String> txtFiles=new HashMap<String, String>();
    static {
        txtFiles.put("当日赔款","当年当日赔款金额.txt");
        txtFiles.put("当日签单","当年当日签单保费.txt");
        txtFiles.put("当日实收","当年当日实收保费.txt");
        txtFiles.put("当月赔款","当年当月赔款金额.txt");
        txtFiles.put("当月签单","当年当月签单保费.txt");
        txtFiles.put("当月实收","当年当月实收保费.txt");
        txtFiles.put("当年赔款","当年累计赔款金额.txt");
        txtFiles.put("当年签单","当年累计签单保费.txt");
        txtFiles.put("当年实收","当年累计实收保费.txt");

        txtFiles.put("去年当日赔款","去年当日赔款金额.txt");
        txtFiles.put("去年当日签单","去年当日签单保费.txt");
        txtFiles.put("去年当日实收","去年当日实收保费.txt");
        txtFiles.put("去年当月赔款","去年当月赔款金额.txt");
        txtFiles.put("去年当月签单","去年当月签单保费.txt");
        txtFiles.put("去年当月实收","去年当月实收保费.txt");
        txtFiles.put("上年赔款","去年累计赔款金额.txt");
        txtFiles.put("上年签单","去年累计签单保费.txt");
        txtFiles.put("上年实收","去年累计实收保费.txt");
    }
    public String basePath=null;
    public String filesPath=null;
    public String templateFilesPath;
    public String targetFilesPath;
    private Date now=new Date();
    public FileScanAndImport(){

    }
    public FileScanAndImport(String basePath,String templateFilesPath,String targetFilesPath){
        this.basePath=basePath;
        this.templateFilesPath=templateFilesPath;
        this.targetFilesPath=targetFilesPath;
    }
    public boolean scanAndImport(){
        boolean result=true;
        if(basePath==null){
            System.out.println("1");
            return false;
        }
        if(filesPath==null||filesPath.trim().equals("")){
            SimpleDateFormat simpleDateFormat=new SimpleDateFormat("yyyyMMdd");
           filesPath=simpleDateFormat.format(now);
        }
        String fullFilesPath=basePath+"/"+filesPath;
        if(basePath.equals("")){
            fullFilesPath=filesPath;
        }
        File file=new File(fullFilesPath);
        System.out.println(fullFilesPath);
        if (!file.exists()||!file.isDirectory()){
            System.out.println("2");
            return false;
        }
        File templateFilesDir=new File(templateFilesPath);
        if(!templateFilesDir.exists()||!templateFilesDir.isDirectory()){
            System.out.println("3");
            return false;
        }
        File targetFilesDir=new File(targetFilesPath);
        if(!targetFilesDir.exists()||!targetFilesDir.isDirectory()){
            targetFilesDir.mkdirs();
        }

        String fileGeneral=templateFileGeneral;
        try{
            SimpleDateFormat simpleDateFormat=new SimpleDateFormat("yyyyMMdd");
            Date date=simpleDateFormat.parse(filesPath);
            Calendar calendar=Calendar.getInstance();
            calendar.setTime(date);
            calendar.add(Calendar.DATE,-1);
            Date newDate=calendar.getTime();
            String dateStr=simpleDateFormat.format(newDate);
            int idx=templateFileGeneral.lastIndexOf(".");
            String pre=templateFileGeneral.substring(0,idx);
            String next=templateFileGeneral.substring(idx,templateFileGeneral.length());
            fileGeneral=pre+dateStr+next;
        }catch (Exception ex){}

        //拷贝模版文件到目标目录
        if(!copyFile(templateFilesPath+"/"+templateFileDetail,targetFilesPath+"/"+templateFileDetail)){
            System.out.println("4");
            return false;
        }
        if(!copyFile(templateFilesPath+"/"+templateFileGeneral,targetFilesPath+"/"+fileGeneral)){
            System.out.println("5");
            return false;
        }

        //导入数据
        ExcelImport excelImport=new ExcelImport(targetFilesPath+"/"+templateFileDetail,null,"yyyy-MM-dd");
        Set keySet=txtFiles.keySet();
        Iterator keyIterator=keySet.iterator();
        while (keyIterator.hasNext()){
            String key=(String)keyIterator.next();
            String value=txtFiles.get(key);
            excelImport.importData(fullFilesPath+"/"+value,key,1,0,null);
            excelImport.cleanNullValue(key,1,11,3265,11);
        }
        excelImport.evaluateAllFormulaCells();

        //计算公式
        ExcelImport excelImportGeneral=new ExcelImport(targetFilesPath+"/"+fileGeneral);
        excelImportGeneral.workbookFullNames=new String[]{targetFilesPath+"/"+templateFileDetail};
        excelImportGeneral.evaluateAllFormulaCells();
        System.out.println("6");
        return result;
    }
    public boolean scanAndImport(String filesPath){
        this.filesPath=filesPath;
        return scanAndImport();
    }
    private boolean copyFile(String filesPath,String targetFilesPath){
        boolean result=true;
        FileOutputStream fileOutputStream=null;
        FileChannel foutChannel=null;
        FileInputStream fileInputStream=null;
        FileChannel finChannel=null;
        try{
            File templateFileDetailFile=new File(filesPath);
            File targetFile=new File(targetFilesPath);
            fileOutputStream=new FileOutputStream(targetFile);
            foutChannel=fileOutputStream.getChannel();
            fileInputStream=new FileInputStream(templateFileDetailFile);
            finChannel=fileInputStream.getChannel();
            finChannel.transferTo(0,finChannel.size(),foutChannel);
        }catch (Exception ex){
            result=false;
        }finally {
            try {
                if(fileInputStream!=null)
                    fileInputStream.close();
                if(finChannel!=null)
                    finChannel.close();
                if(fileOutputStream!=null)
                    fileOutputStream.close();
                if(foutChannel!=null)
                    foutChannel.close();
            }catch (Exception ex){

            }
        }
        return result;
    }
}
