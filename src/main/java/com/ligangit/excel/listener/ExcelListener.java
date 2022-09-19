package com.ligangit.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author daniel
 * @createTime 2022/9/16 0016 11:04
 * @description 解析监听器
 *   每解析一行会回调invoke()方法。
 *   整个excel解析结束会执行doAfterAllAnalysed()方法
 *
 *   没有考虑合并单元格的情况
 */

public class ExcelListener extends AnalysisEventListener<Object> {

    //定义一个保存Excel所有记录的集合
    private List<Object> datas = new ArrayList<>();

    public List<Object> getDatas() {
        return datas;
    }

    public void setDatas(List<Object> datas) {
        this.datas = datas;
    }

    /**
     * 逐行解析
     * object : 当前行的数据
     * 这个每一条数据解析都会来调用
     * 我们将每一条数据都保存到list集合中
     * @param object    one row value. Is is same as {@link AnalysisContext#readRowHolder()}
     * @param analysisContext
     */
    @Override
    public void invoke(Object object, AnalysisContext analysisContext) {
        System.out.println("读取object=" + object);
        //当前行
        // context.getCurrentRowNum()
        //数据存储到list，供批量处理，或后续自己业务逻辑处理。
        if (object != null) {
            datas.add(object);
//            handleBusinessLogic();
        }

          /*
        如数据过大，可以进行定量分批处理
        if(datas.size() >= 200){
            handleBusinessLogic();
            datas.clear();
        }
         */
    }

    /**
     * 读取表头内容
     * @param headMap 表头
     * @param analysisContext
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext analysisContext) {
        System.out.println("表头" + headMap);
    }

    /**
     * 解析完所有数据后会调用该方法
     * @param analysisContext
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        //解析结束销毁不用的资源,非必要语句，查看导入的数据
        System.out.println("读取Excel完毕" + datas.size());
    }

    //根据业务自行实现该方法，例如将解析好的dataList存储到数据库中
    private void handleBusinessLogic() {}
}
