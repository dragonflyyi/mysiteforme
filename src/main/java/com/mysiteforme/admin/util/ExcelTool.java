package com.mysiteforme.admin.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;

import java.io.*;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

/**
 * Created by dragonfly on 2018/10/21.
 */
public class ExcelTool {
    /**
     * 最大导入行数
     */
    public static int  MAX_IMPORT_ROW=2000;
    /**
     * 最大导出行数
     */
    public static int  MAX_EXPORT_ROW=3000;


    public static void  list2File(List<?> list, Class cls, String exportFileName) throws ExcelException {

        if(list.size()==0 || list==null){
            throw new ExcelException("数据源中没有任何数据");
        }


         Field[]  fields = cls.getDeclaredFields();
         if(fields==null || fields.length <=0 ){
             throw new ExcelException("获取导出类异常");
         }

         try{
             HSSFWorkbook workbook  = new HSSFWorkbook();
             HSSFSheet sheet = workbook.createSheet();
             for(int row=0; row <list.size(); row++){
                 HSSFRow rowObj =  sheet.createRow(row);
                 Map<String, Object>  beanMap = ToolUtil.convertBeanToMap(list.get(row));

                 for(int index=0; index <fields.length; index++){
                     HSSFCell cell = rowObj.createCell(index);
                     if(row > 0){
                         cell.setCellValue(""+beanMap.get(fields[index].getName()));
                     }else{
                         cell.setCellValue(""+fields[index].getName());
                     }




                 }
             }

             FileOutputStream output=new FileOutputStream(exportFileName);
             workbook.write(output);
             output.flush();
             output.close();


         }catch (Exception e){
            e.printStackTrace();

         }finally {

         }

    }




}
