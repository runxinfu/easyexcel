package com.uton.excel.support;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.uton.excel.metadata.BaseRowModel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * @author: furunxin
 * @Date: 2018/11/16
 * @Description:
 */
public class EasyExcelWriterFactroy extends ExcelWriter {
    private OutputStream outputStream;
    private int sheetNo = 1;

    public EasyExcelWriterFactroy(OutputStream outputStream, ExcelTypeEnum typeEnum) {
        super(outputStream, typeEnum);
        this.outputStream = outputStream;
    }

    public EasyExcelWriterFactroy write(List<? extends BaseRowModel> list, String sheetName, BaseRowModel object) {
        ++this.sheetNo;
        try {
            Sheet sheet = new Sheet(this.sheetNo, 0,object.getClass());
            sheet.setSheetName(sheetName);
            this.write(list, sheet);
        } catch (Exception exception) {
            exception.printStackTrace();
            try {
                this.outputStream.flush();
            } catch (IOException ioException) {
                ioException.printStackTrace();
            }
        }
        return this;
    }

    public EasyExcelWriterFactroy write0(List<List<String>> list, String sheetName, Table table) {
        ++this.sheetNo;
        try {
            Sheet sheet = new Sheet(this.sheetNo,0);
            sheet.setSheetName(sheetName);
            this.write0(list,sheet,table);
        } catch (Exception exception) {
            exception.printStackTrace();
            try {
                this.outputStream.flush();
            } catch (IOException ioException) {
                ioException.printStackTrace();
            }
        }
        return this;
    }


    @Override
    public void finish() {
        super.finish();
        try {
            this.outputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
