package com.ligangit.excel.controller;


import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.ligangit.excel.config.CommonCellStyleStrategy;
import com.ligangit.excel.entity.Member;
import com.ligangit.excel.handler.CustomCellWriteHandler;
import com.ligangit.excel.handler.ExcelLockingHandler;
import com.ligangit.excel.handler.WaterMarkHandler;
import com.ligangit.excel.util.ExcelUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author daniel
 * @createTime 2022/9/16 0016 15:44
 * @description
 */

@Slf4j
@Controller
@RequestMapping("write")
public class WriteController {

    /**
     * 普通导出方式,不使用工具类
     */
    @PostMapping("/export1")
    public void exportMembers1(HttpServletResponse response, HttpServletRequest request) throws IOException {
        //从数据库获取数据
        //List<Member> members = memberService.getAllMember();
        try {
            String filenames = "demo";
            String userAgent = request.getHeader("User-Agent");
            if (userAgent.contains("MSIE") || userAgent.contains("Trident")) {
                filenames = URLEncoder.encode(filenames, "UTF-8");
            } else {
                filenames = new String(filenames.getBytes("UTF-8"), "ISO-8859-1");
            }
            response.setContentType("application/vnd.ms-exce");
            response.setCharacterEncoding("utf-8");
            response.addHeader("Content-Disposition", "attachment;filename=" + filenames + ".xlsx");
            //从数据库获取数据，更换data()即可
            EasyExcel.write(response.getOutputStream(), Member.class)
                    // EasyExcel使用时需要设置inMemory(true)，否则默认使用的是SXSSFWorkbook，会报错！
                    .inMemory(true)
                    // 设置水印
                    .registerWriteHandler(new WaterMarkHandler("这是水印文案内容"))
                    .sheet("sheet").doWrite(data());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 基于策略及拦截器导出(带样式)
     */
    @RequestMapping("/export2")
    public void exportMembers2(HttpServletResponse response, HttpServletRequest request) throws IOException {
//        List<Member> members = memberService.getAllMember();
        try {
            String filenames = "demo2";
            String userAgent = request.getHeader("User-Agent");
            if (userAgent.contains("MSIE") || userAgent.contains("Trident")) {
                filenames = URLEncoder.encode(filenames, "UTF-8");
            } else {
                filenames = new String(filenames.getBytes("UTF-8"), "ISO-8859-1");
            }
            response.setContentType("application/vnd.ms-exce");
            response.setCharacterEncoding("utf-8");
            response.addHeader("Content-Disposition", "attachment;filename=" + filenames + ".xlsx");

            EasyExcel.write(response.getOutputStream(), Member.class)
//                    .excelType(ExcelTypeEnum.XLS)
                    // EasyExcel使用时需要设置inMemory(true)，否则默认使用的是SXSSFWorkbook，会报错！
                    .inMemory(true)
                    .password("123456")
                    // 设置水印
                    .registerWriteHandler(new WaterMarkHandler("这是水印文案内容"))
                    // 设置只读
                    .registerWriteHandler(new ExcelLockingHandler())
                    // 注册通用格式策略
//                    .registerWriteHandler(CommonCellStyleStrategy.getHorizontalCellStyleStrategy())
                    // 设置自定义格式策略
//                .registerWriteHandler(new CustomCellWriteHandler())
                    .sheet("sheet")
                    .doWrite(data());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 封装工具类导出方式
     */
    @PostMapping("/export3")
    @ResponseStatus(HttpStatus.CREATED)
    public void inputExcel(@RequestBody MultipartFile file) {

        //工具类方法1:
//        ExcelUtils.getExcelimporttemplate(Member.class,"test1",data());

        //工具类方法2:
//        ExcelUtils.excelExport(Member.class,"test1",data());

        //工具类方法3:
        List<String> list = new ArrayList<>();
        //列名
        list.add("111");
        list.add("222");
        list.add("333");
        ExcelUtils.getBigTitleExcel("测试big",Member.class,"test3",data(),list);
    }

    //辅助方法
    private List<Member> data() {
        List<Member> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Member data = new Member();
            data.setUsername("Helen" + i);
            data.setBirthday(new Date());
            if (i%2 == 0) {
                data.setGender(1);
            }else {
                data.setGender(0);
            }
            list.add(data);
        }
        return list;
    }

}

