package com.uton.excel;

import com.alibaba.excel.support.ExcelTypeEnum;
import com.uton.excel.metadata.BaseRowModel;
import com.uton.excel.support.EasyExcelException;
import com.uton.excel.support.EasyExcelListener;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.util.List;

/**
 * @author: furunxin
 * @Date: 2018/11/17 19:26
 * @Description:
 */
public class ReaderExcel {

    /**
      * @description 读取Excel(允许多个sheet)
      * @Params [excel, object, headLineMun]
      * exlce:需要读取的excle文件
      * object:继承BaseRowModel的实体,可传null,传null则返回List<List<String>>
      * headLineMun:从第几行开始读取Excel,主要作用作用就是过滤Excel头部标题
      * @return java.util.List<java.lang.Object>
      * @Author furunxin
      * @Date 2018/11/20 上午10:23
      **/
    public static List<Object> reader(File excel, BaseRowModel object,int headLineMun){
        return getReader((File)excel, object, 0,headLineMun);
    }

    /**
      * @description 根据指定的sheetNo来读取Excel
      * @Params [excel, object, sheetNo, headLineMun]
      * exlce:需要读取的excle文件
      * object:继承BaseRowModel的实体,可传null,传null则返回List<List<String>>
      * sheetNo:指定读取那个sheet页
      * XLS 类型文件 sheet 序号为顺序，第一个 sheet 序号为 1
      * XLSX 类型 sheet 序号顺序为倒序，即最后一个 sheet 序号为 1
      * headLineMun:从第几行开始读取Excel,主要作用作用就是过滤Excel头部标题
      * @return java.util.List<java.lang.Object>
      * @Author furunxin
      * @Date 2018/11/20 上午10:27
      **/
    public static List<Object> reader(File excel, BaseRowModel object, int sheetNo,int headLineMun) {
        return getReader(excel, object, sheetNo,headLineMun);
    }

    /**
      * @description 读取远程Excel文件(允许多个sheet)
      * @Params [fileUrl, fileName, object, headLineMun]
      * fileUrl:文件地址
      * fileName:文件名
      * object:继承BaseRowModel的实体,可传null,传null则返回List<List<String>>
      * headLineMun:从第几行开始读取Excel,主要作用作用就是过滤Excel头部标题
      * @return java.util.List<java.lang.Object>
      * @Author furunxin
      * @Date 2018/11/20 上午10:28
      **/
    public static List<Object> reader0(String fileUrl,String fileName,BaseRowModel object,int headLineMun){
        InputStream inputStream = null;
        try {
            inputStream = urlConvertInputStream(fileUrl);
        } catch (IOException e) {
            throw new EasyExcelException("文件读取失败");
        }
        return getReader(fileName, inputStream, object, 0,headLineMun);
    }

    /**
      * @description 读取远程Excel文件(指定sheetNo)
      * @Params [fileUrl, fileName, object, sheetNo, headLineMun]
      * fileUrl:文件地址
      * fileName:文件名
      * object:继承BaseRowModel的实体,可传null,传null则返回List<List<String>>
      * sheetNo:指定读取那个sheet页
      * XLS 类型文件 sheet 序号为顺序，第一个 sheet 序号为 1
      * XLSX 类型 sheet 序号顺序为倒序，即最后一个 sheet 序号为 1
      * headLineMun:从第几行开始读取Excel,主要作用作用就是过滤Excel头部标题
      * @return java.util.List<java.lang.Object>
      * @Author furunxin
      * @Date 2018/11/20 上午10:55
      **/
    public static List<Object> reader0(String fileUrl,String fileName,BaseRowModel object, int sheetNo,int headLineMun){
        InputStream inputStream = null;
        try {
            inputStream = urlConvertInputStream(fileUrl);
        } catch (IOException e) {
            throw new EasyExcelException("文件流读取失败");
        }
        return getReader(fileName, inputStream, object, sheetNo,headLineMun);
    }


    private static List<Object> getReader(File excel, BaseRowModel object, int sheetNo,int headLineMun) {
        String fileName = excel.getName();

        try {
            InputStream inputStream = new FileInputStream(excel);
            return getReader(fileName, inputStream, object, sheetNo,headLineMun);
        } catch (IOException var6) {
            var6.printStackTrace();
            return null;
        }
    }

    private static List<Object> getReader(String fileName, InputStream inputStream, BaseRowModel object, int sheetNo,int headLineMun) {
        if (fileName == null || !fileName.toLowerCase().endsWith(".xls") && !fileName.toLowerCase().endsWith(".xlsx")) {
            throw new EasyExcelException("文件格式错误！");
        } else {
            ExcelTypeEnum excelTypeEnum = ExcelTypeEnum.XLSX;
            if (fileName.toLowerCase().endsWith(".xls")) {
                excelTypeEnum = ExcelTypeEnum.XLS;
            }

            EasyExcelListener easyExcelListener = new EasyExcelListener();
            com.alibaba.excel.ExcelReader reader = new com.alibaba.excel.ExcelReader(inputStream,(Object) null, easyExcelListener);
            if (reader == null) {
                return null;
            } else {
                if (sheetNo == 0){
                    List<com.alibaba.excel.metadata.Sheet> sheets = reader.getSheets();
                    for (com.alibaba.excel.metadata.Sheet sheet : sheets){
                        if (object != null){
                            sheet.setClazz(object.getClass());
                        }
                        sheet.setHeadLineMun(headLineMun);
                        reader.read(sheet);
                    }
                }else{
                    com.alibaba.excel.metadata.Sheet sheet = new com.alibaba.excel.metadata.Sheet(sheetNo);
                    if (object != null) {
                        sheet.setClazz(object.getClass());
                    }
                    sheet.setHeadLineMun(headLineMun);
                    reader.read(sheet);
                }
                return easyExcelListener.getDatas();
            }
        }
    }

    private static InputStream urlConvertInputStream(String fileUrl) throws IOException {
        URL url = new URL(fileUrl);
        URLConnection con = url.openConnection();
        con.setConnectTimeout(5*1000);
        InputStream inputStream = con.getInputStream();
        return inputStream;
    }

}
