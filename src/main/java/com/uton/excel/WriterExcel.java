package com.uton.excel;

import com.alibaba.excel.support.ExcelTypeEnum;
import com.uton.excel.metadata.BaseRowModel;
import com.uton.excel.metadata.Table;
import com.uton.excel.support.EasyExcelException;
import com.uton.excel.support.EasyExcelWriterFactroy;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author: furunxin
 * @Date: 2018/11/17 19:42
 * @Description:
 */
public class WriterExcel {

    /**
      * @description excel导出到指定文件路径
      * @Params [data, fields, filePath, sheetName]
      * data:list<Map>对象
      * fields:标题数组
      * filePath:excel文件保存路径
      * sheetName:excel SheetName
      * @return void
      * @Author furunxin
      * @Date 2018/11/20 上午10:57
      **/
    public static void writer2Map(List<Map<Object,Object>> data,String[] fields, String filePath, String sheetName){
        try {
            List<List<String>> writeData = convertMapData(data);
            Table table = converWriteHeader(fields);
            OutputStream out = new FileOutputStream(filePath);
            EasyExcelWriterFactroy writerFactroy = new EasyExcelWriterFactroy(out, ExcelTypeEnum.XLSX);
            writerFactroy.write0(writeData,sheetName,table).finish();
        } catch (FileNotFoundException e) {
            throw new EasyExcelException("文件路径不存在");
        }
    }

    /**
     * @description excel导出到指定文件路径
     * @Params [data, fields, filePath, sheetName]
     * data:Object对象
     * fields:标题数组
     * filePath:excel文件保存路径
     * sheetName:excel SheetName
     * @return void
     * @Author furunxin
     * @Date 2018/11/20 上午10:57
     **/
    public static void writer2Object(List<Object> data,String[] fields, String filePath, String sheetName){
        try {
            List<List<String>> writeData = convertObjectData(data);
            Table table = converWriteHeader(fields);
            OutputStream out = new FileOutputStream(filePath);
            EasyExcelWriterFactroy writerFactroy = new EasyExcelWriterFactroy(out,ExcelTypeEnum.XLSX);
            writerFactroy.write0(writeData,sheetName,table).finish();
        } catch (FileNotFoundException e) {
            throw new EasyExcelException("文件路径不存在");
        } catch (IllegalAccessException e) {
            throw new EasyExcelException(e.getMessage());
        }
    }

    /**
      * @description excel导出到指定文件路径
      * @Params [data, filePath, sheetName, object]
      * data: ? extends BaseRowModel list
      * filePath:excel文件保存路径
      * sheetName:excel SheetName
      * object:? extends BaseRowModel
       * @return void
      * @Author furunxin
      * @Date 2018/11/20 上午11:00
      **/
    public static void writer2BaseRowModel(List<? extends BaseRowModel> data, String filePath, String sheetName, BaseRowModel object){
        try{
            OutputStream out = new FileOutputStream(filePath);
            EasyExcelWriterFactroy writerFactroy = new EasyExcelWriterFactroy(out,ExcelTypeEnum.XLSX);
            writerFactroy.write(data,sheetName,object).finish();
        }catch (FileNotFoundException e){
            throw new EasyExcelException("文件路径不存在");
        }
    }

    /**
      * @description 基于BaseRowModel进行客户端Excel导出
      * @Params [response, data, fileName, sheetName, object]
      * @return void
      * @Author furunxin
      * @Date 2018/11/20 上午11:03
      **/
    public static void writer2BaseRowModel(HttpServletResponse response,List<? extends BaseRowModel> data, String fileName, String sheetName, BaseRowModel object){
        try{
            OutputStream out = getOutputStream(response, fileName);
            EasyExcelWriterFactroy writerFactroy = new EasyExcelWriterFactroy(out, ExcelTypeEnum.XLSX);
            writerFactroy.write(data,sheetName,object).finish();
        } catch (IOException e) {
            throw new EasyExcelException("文件写入失败");
        }
    }

    /**
      * @description 客户端Excel导出
      * @Params [response, data, fields, fileName, sheetName]
      * @return void
      * @Author furunxin
      * @Date 2018/11/20 上午11:04
      **/
    public static void writer2Map(HttpServletResponse response, List<Map<Object,Object>> data, String[] fields,String fileName, String sheetName){
        try {

            OutputStream out = getOutputStream(response, fileName);
            EasyExcelWriterFactroy writerFactroy = new EasyExcelWriterFactroy(out, ExcelTypeEnum.XLSX);
            List<List<String>> writeData = convertMapData(data);
            Table table = converWriteHeader(fields);
            writerFactroy.write0(writeData,sheetName,table).finish();
        } catch (IOException e) {
            throw new EasyExcelException("文件流写入失败");
        }
    }

    /**
      * @description 客户端Excel导出
      * @Params [response, data, fields, fileName, sheetName]
      * @return void
      * @Author furunxin
      * @Date 2018/11/20 上午11:04
      **/
    public static void writer2Object(HttpServletResponse response, List<Object> data, String[] fields,String fileName, String sheetName){
        try {
            OutputStream out = getOutputStream(response, fileName);
            EasyExcelWriterFactroy writerFactroy = new EasyExcelWriterFactroy(out, ExcelTypeEnum.XLSX);
            List<List<String>> writeData = convertObjectData(data);
            Table table = converWriteHeader(fields);
            writerFactroy.write0(writeData,sheetName,table).finish();
        } catch (IOException e) {
            throw new EasyExcelException("文件流写入失败");
        } catch (IllegalAccessException e) {
            throw new EasyExcelException(e.getMessage());
        }
    }


    private static List<List<String>> convertMapData(List<Map<Object,Object>> maps){
        List<List<String>> data = new ArrayList();
        Iterator var2 = maps.iterator();
        while(var2.hasNext()) {
            Map<Object, Object> map = (Map)var2.next();
            List<String> stringList = new ArrayList();
            Iterator var5 = map.entrySet().iterator();
            while(var5.hasNext()) {
                Map.Entry<Object, Object> entry = (Map.Entry)var5.next();
                stringList.add(String.valueOf(entry.getValue()));
            }
            data.add(stringList);
        }
        return data;
    }

    private static List<List<String>> convertObjectData(List<Object> objects) throws IllegalAccessException {
        List<List<String>> data = new ArrayList();
        Iterator var2 = objects.iterator();
        while(var2.hasNext()) {
            Object object = var2.next();
            List<String> stringList = new ArrayList();
            Class objectClass = object.getClass();
            Field[] fields = objectClass.getDeclaredFields();

            for(int i = 0; i < fields.length; ++i) {
                Field field = fields[i];
                field.setAccessible(true);
                Object val = field.get(object);
                stringList.add(String.valueOf(val));
            }
            data.add(stringList);
        }
        return data;
    }

    private static Table converWriteHeader(String[] fields){
        List<List<String>> header = new ArrayList<>();
        for (int i= 0;i<fields.length;i++){
            List<String> list = new ArrayList<>();
            list.add(fields[i]);
            header.add(list);
        }
        Table table = new Table(1);
        table.setHead(header);
        return table;
    }

    private static OutputStream getOutputStream(HttpServletResponse response,String fileName) throws IOException {
        String filePath = fileName + ".xlsx";
        File dbfFile = new File(filePath);
        if (!dbfFile.exists() || dbfFile.isDirectory()) {
            dbfFile.createNewFile();
        }
        fileName = new String(filePath.getBytes(), "ISO-8859-1");
        response.addHeader("Content-Disposition", "filename=" + fileName);
        OutputStream out = response.getOutputStream();
        return out;
    }
}
