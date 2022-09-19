package com.ligangit.excel.util;

import cn.hutool.core.convert.Convert;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.ligangit.excel.handler.WaterMarkHandler;
import com.ligangit.excel.listener.ExcelListener;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

/**
 * @author daniel
 * @createTime 2022/9/16 0016 9:47
 * @description
 */
public class ExcelUtils {

    private static final Logger log = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 读取Excel（多个sheet可以用同一个实体类解析）
     * @param excelInputStream
     * @param fileName
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readExcel(InputStream excelInputStream, String fileName, Class<T> clazz) {
        ExcelListener excelListener = new ExcelListener();
        ExcelReader excelReader = getReader(excelInputStream, fileName,clazz, excelListener);
        if (excelReader == null) {
            return new ArrayList<>();
        }
        List<ReadSheet> readSheetList = excelReader.excelExecutor().sheetList();
        for (ReadSheet readSheet : readSheetList) {
            excelReader.read(readSheet);
        }
        excelReader.finish();
        return Convert.toList(clazz, excelListener.getDatas());
    }

    /**
     * 导出数据
     *
     * @param head      类名
     * @param excelname excel名字
     * @param data      数据
     */
    public static void getExcelimporttemplate(Class head, String excelname, List data) {
        ServletRequestAttributes requestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = requestAttributes.getResponse();
        // 这里注意 有人反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        //response.setContentType("application/vnd.ms-excel");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");

        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        try {
            //表格头部样式
            WriteCellStyle headWriteCellStyle = new WriteCellStyle();
            //设置背景颜色
            //headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            //设置头字体
            //WriteFont headWriteFont = new WriteFont();
            //headWriteFont.setFontHeightInPoints((short)13);
            //headWriteFont.setBold(true);
            //headWriteCellStyle.setWriteFont(headWriteFont);
            //设置头居中
            headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);

            //表格内容样式
            WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
            //设置 水平居中
            contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);

            HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
            String fileName = URLEncoder.encode(excelname, "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            EasyExcel.write(response.getOutputStream(), head).registerWriteHandler(horizontalCellStyleStrategy).sheet("Sheet1").doWrite(data);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 导出数据
     *
     * @param head      类名
     * @param excelname excel名字
     * @param data      数据
     */
    public static void excelExport(Class head, String excelname, List data) {
        ServletRequestAttributes requestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = requestAttributes.getResponse();
        // 这里注意 有人反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        //response.setContentType("application/vnd.ms-excel");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        try {
            String fileName = URLEncoder.encode(excelname, "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            EasyExcel.write(response.getOutputStream(), head).sheet("Sheet1").doWrite(data);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 导出数据二级表头
     * (合并第一行单元格,第二行开始是业务字段)
     * @param bigTitle  合并单元格的名字
     * @param head      类名
     * @param excelname 导出的excel名字
     * @param data      数据
     * @param header    导出的字段列名
     */
    public static void getBigTitleExcel(String bigTitle, Class head, String excelname, List data, List<String> header) {
        ServletRequestAttributes requestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = requestAttributes.getResponse();
        // 这里注意 有人反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        //response.setContentType("application/vnd.ms-excel");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
        response.setCharacterEncoding("utf-8");

        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        try {
            //表格头部样式
            WriteCellStyle headWriteCellStyle = new WriteCellStyle();
            //设置背景颜色
            //headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            //设置头字体
            //WriteFont headWriteFont = new WriteFont();
            //headWriteFont.setFontHeightInPoints((short)13);
            //headWriteFont.setBold(true);
            //headWriteCellStyle.setWriteFont(headWriteFont);
            //设置头居中
            headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);

            //表格内容样式
            WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
            //设置 水平居中
            contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);

            HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
            String fileName = URLEncoder.encode(excelname, "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            EasyExcel.write(response.getOutputStream())
                    .head(getExcelHeader(header, bigTitle))
                    // EasyExcel使用时需要设置inMemory(true)，否则默认使用的是SXSSFWorkbook，会报错！
                    .inMemory(true)
                    .registerWriteHandler(new WaterMarkHandler("wwwwwwwwwwwwwwww"))
                    .registerWriteHandler(horizontalCellStyleStrategy).sheet("Sheet1")
                    .doWrite(data);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private static List<List<String>> getExcelHeader(List<String> header, String bigTitle) {
        List<List<String>> head = new ArrayList<>();
        List<String> head0 = null;
        for (String h : header) {
            head0 = new ArrayList<>();
            head0.add(bigTitle);
            head0.add(h);
            head.add(head0);
        }
        return head;
    }

    /**
     * 返回ExcelReader
     *
     * @param excel         文件
     * @param clazz         实体类
     * @param excelListener
     */
    private static <T> ExcelReader getReader(InputStream inputStream, String filename, Class<T> clazz, ExcelListener excelListener) {
        try {
            if (filename == null ||
                    (!filename.toLowerCase().endsWith(".xls") && !filename.toLowerCase().endsWith(".xlsx"))) {
                return null;
            }
            ExcelReader excelReader = EasyExcel.read(inputStream, clazz, excelListener).build();
            inputStream.close();
            return excelReader;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}
