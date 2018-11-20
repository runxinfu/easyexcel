package com.uton.excel.support;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;

/**
 * @author: furunxin
 * @Date: 2018/11/16
 * @Description:
 */
public class EasyExcelListener extends AnalysisEventListener {

    private List<Object> datas = new ArrayList();

    public EasyExcelListener() {}

    @Override
    public void invoke(Object o, AnalysisContext analysisContext) {
        this.datas.add(o);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }

    public List<Object> getDatas() {
        return this.datas;
    }

    public void setDatas(List<Object> datas) {
        this.datas = datas;
    }
}
