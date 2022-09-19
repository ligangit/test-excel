package com.ligangit.excel.controller;

import com.alibaba.excel.EasyExcel;
import com.ligangit.excel.entity.Member;
import com.ligangit.excel.listener.ExcelListener;
import com.ligangit.excel.util.ExcelUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

/**
 * @author daniel
 * @createTime 2022/9/16 0016 15:39
 * @description easyexcel 几种读取文件的方式
 *      方式一：同步读取，将解析结果返回，比如返回List<Member>，业务再进行相应的数据集中处理
 *      方式二：对照doReadSync()方法的是最后调用doRead()方法，不进行结果返回，而是在MemberImportExcelListener中进行一条条数据的处理；
 */

@Controller
@RequestMapping("read")
public class ReadController {

    /**
     * 1. 从Excel导入内容
     * 基于同步获取结果列表的形式进行导入
     */
    @PostMapping("/import1")
    @ResponseBody
    public void importMemberList(@RequestPart("file") MultipartFile file) throws IOException {
        List<Member> list = EasyExcel.read(file.getInputStream())
                .head(Member.class)
                .sheet()
                .doReadSync();
        //打印读取的每条数据
        for (Member member : list) {
            System.out.println(member);
        }
    }

    /**
     * 2. 基于Listener方式从Excel导入会员列表
     */
    @RequestMapping(value = "/import2", method = RequestMethod.POST)
    @ResponseBody
    public void importMemberList2(@RequestPart("file") MultipartFile file) throws IOException {
        // 此处示例为方式二,读取监听器
        EasyExcel.read(file.getInputStream(), Member.class, new ExcelListener()).sheet().doRead();
    }

    /**
     *  基于ExcelUtils工具类导入
     */
    @PostMapping(value = "/import3")
    @ResponseBody
    public List<Member> read(MultipartFile excel) throws IOException {
        return ExcelUtils.readExcel(excel.getInputStream(), excel.getOriginalFilename(), Member.class);
    }
}

