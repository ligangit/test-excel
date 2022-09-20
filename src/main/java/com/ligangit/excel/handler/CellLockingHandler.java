package com.ligangit.excel.handler;

//import com.alibaba.excel.metadata.CellData;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;

/**
 * @author daniel
 * @createTime 2022/9/16 0016 10:44
 * @description 实现CellWriteHandler接口, 实现对单元格样式的精确控制
 */
public class CellLockingHandler implements CellWriteHandler {

    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer integer, Integer integer1, Boolean aBoolean) {

    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer integer, Boolean aBoolean) {

    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> list, Cell cell, Head head, Integer integer, Boolean aBoolean) {
//        //可以写的模式
//        CellStyle unLockCellStyle = cell.getSheet().getWorkbook().createCellStyle();
//        unLockCellStyle.setLocked(false);
//        //默认只读，设置非第一列可以编辑,无数据的表格不会创建cell，此方式无效
//        if (cell.getColumnIndex() != 0) {
//            cell.setCellStyle(unLockCellStyle);
//        }
        //只读模式
        CellStyle lockedCellStyle = cell.getSheet().getWorkbook().createCellStyle();
        lockedCellStyle.setLocked(true);
        //可以写的模式
        CellStyle unLockCellStyle = cell.getSheet().getWorkbook().createCellStyle();
        unLockCellStyle.setLocked(false);
        //对第二列进行只读模式
        if (cell.getColumnIndex() == 1) {

            cell.setCellStyle(lockedCellStyle);
        } else {
            cell.setCellStyle(unLockCellStyle);
        }
        //设置表格的保护 只有设置表格保护， cellStyle locked 模式才有效果
        //https://my.oschina.net/u/4399738/blog/3700765(这篇博客有介绍)
        // 因为我在ExcelLockingHandler中处理了，所以注释掉了
//        cell.getSheet().protectSheet("123");

    }
}

