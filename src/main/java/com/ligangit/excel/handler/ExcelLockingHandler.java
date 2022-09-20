package com.ligangit.excel.handler;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * @author daniel
 * @createTime 2022/9/16 0030 16:37
 * @description
 * 		excel设置只读，只支持XSSFWorkbook，其余类别：SXSSFWorkbook、SXSSFWorkbook请另寻他法
 *
 * easyExcel使用时需要设置inMemory(true)，否者默认使用的是SXSSFWorkbook，会报错！
 *
 *  打印水印
 */
@RequiredArgsConstructor
public class ExcelLockingHandler implements SheetWriteHandler {

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
    }

    @SneakyThrows
    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        XSSFSheet sheet = (XSSFSheet) writeSheetHolder.getSheet();

        //设置只读
        sheet.enableLocking();
        // 如果需要定义比较详细的权限，需要设置如下内容
        // 定义比较详细的excel权限 ------开始
        CTSheetProtection sheetProtection = sheet.getCTWorksheet().getSheetProtection();
        sheetProtection.setSelectLockedCells(false);
        sheetProtection.setSelectUnlockedCells(false);
        sheetProtection.setFormatCells(false);
        sheetProtection.setFormatColumns(true);
        sheetProtection.setFormatRows(true);
        sheetProtection.setInsertColumns(true);
        sheetProtection.setInsertRows(true);
        sheetProtection.setInsertHyperlinks(true);
        sheetProtection.setDeleteColumns(true);
        sheetProtection.setDeleteRows(true);
        sheetProtection.setSort(true);
        sheetProtection.setAutoFilter(true);
        sheetProtection.setPivotTables(true);
        sheetProtection.setObjects(true);
        sheetProtection.setScenarios(true);
        // 定义比较详细的excel权限 ------结束
        //333是密码
        sheet.protectSheet(new String("333"));
    }
}

