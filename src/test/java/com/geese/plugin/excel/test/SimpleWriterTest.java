package com.geese.plugin.excel.test;

import com.geese.plugin.excel.SimpleWriter;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * SimpleWriter 接口测试
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 18:06
 * @sine 0.0.1
 */
public class SimpleWriterTest {
    private static OutputStream output;
    private static InputStream template;

    @BeforeClass
    public static void beforeClass() throws IOException {
        // 输出
        URL url = Thread.currentThread().getContextClassLoader().getResource("demo-writer.xlsx");
        output = new FileOutputStream(url.getFile());

        // 模板
        url = Thread.currentThread().getContextClassLoader().getResource("demo-writer-template.xlsx");
        template = new FileInputStream(url.getFile());
    }

    /**
     * 实例1：快速使用
     */
    @Test
    public void test001() {
        // ExcelMapping 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
        // 准备数据, 每一行是一个Map, 每一个表格是一个List<Map>
        String names = "鲁沛儿 鲁天薇 鲁飞雨 鲁天纵 鲁白梦 鲁嘉胜 鲁盼巧 鲁访天 鲁清妍 鲁盼晴 张馨蓉 张白萱 张若云 张雅畅 张雅寒 张雨华";
        List<Map> tableData = new ArrayList<>();
        Map rowData;
        for (String name : names.split("\\s+")) {
            rowData = new HashMap();
            rowData.put("name", name);
            rowData.put("age", Double.valueOf(Math.random() * 100).intValue());
            rowData.put("idCard", Double.valueOf(Math.random() * 1000000000).longValue());
            rowData.put("qq", Double.valueOf(Math.random() * 1000000000).longValue());
            rowData.put("email", Double.valueOf(Math.random() * 1000000000).longValue() + "@163.com");
            rowData.put("phone", Double.valueOf(Math.random() * 1000000000).longValue());
            tableData.add(rowData);
        }
        // 通过SimpleWriter类操作
        SimpleWriter.build(output)  // 必选，将生成的excel输出到什么地方
                .insert("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")  // 必选，数据Key与Excel列的映射
                .addData(tableData) // 必选，插入的数据
                .execute(); // 执行
    }

    /**
     * 实例2：使用可选配置
     */
    @Test
    public void test002() {
        // ExcelMapping 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
        // 准备数据, 每一行是一个Map, 每一个表格是一个List<Map>
        String names = "鲁沛儿 鲁天薇 鲁飞雨 鲁天纵 鲁白梦 鲁嘉胜 鲁盼巧 鲁访天 鲁清妍 鲁盼晴 张馨蓉 张白萱 张若云 张雅畅 张雅寒 张雨华";
        List<Map> tableData = new ArrayList<>();
        Map rowData;
        for (String name : names.split("\\s+")) {
            rowData = new HashMap();
            rowData.put("name", name);
            rowData.put("age", Double.valueOf(Math.random() * 100).intValue());
            rowData.put("idCard", Double.valueOf(Math.random() * 1000000000).longValue());
            rowData.put("qq", Double.valueOf(Math.random() * 1000000000).longValue());
            rowData.put("email", Double.valueOf(Math.random() * 1000000000).longValue() + "@163.com");
            rowData.put("phone", Double.valueOf(Math.random() * 1000000000).longValue());
            tableData.add(rowData);
        }
        // 通过SimpleWriter类操作
        SimpleWriter.build(output)  // 必选，将生成的excel输出到什么地方
                .insert("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")  // 必选，数据Key与Excel列的映射
                .into("0xx")  // 可选（默认：插入到第0个sheet, 可以指定名称，比如：xx数据报表）
                .limit(0, 10) // 可选（参数1：从哪行开始插入，参数2：插入多少航，默认：0,tableData.size()）
                .addData(tableData) // 必选，插入的数据
                .execute(); // 执行
    }

    /**
     * 实例4：使用模板
     */
    @Test
    public void test004() {
        // ExcelMapping 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
        // 准备数据, 每一行是一个Map, 每一个表格是一个List<Map>
        String names = "鲁沛儿 鲁天薇 鲁飞雨 鲁天纵 鲁白梦 鲁嘉胜 鲁盼巧 鲁访天 鲁清妍 鲁盼晴 张馨蓉 张白萱 张若云 张雅畅 张雅寒 张雨华";
        List<Map> tableData = new ArrayList<>();
        Map rowData;
        for (String name : names.split("\\s+")) {
            rowData = new HashMap();
            rowData.put("name", name);
            rowData.put("age", Double.valueOf(Math.random() * 100).intValue());
            rowData.put("idCard", Double.valueOf(Math.random() * 1000000000).longValue());
            rowData.put("qq", Double.valueOf(Math.random() * 1000000000).longValue());
            rowData.put("email", Double.valueOf(Math.random() * 1000000000).longValue() + "@163.com");
            rowData.put("phone", Double.valueOf(Math.random() * 1000000000).longValue());
            tableData.add(rowData);
        }
        // 通过SimpleWriter类操作
        SimpleWriter.build(output, template)  // 必选，将按照模板去生成excel
                .insert("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")  // 必选，数据Key与Excel列的映射
                .limit(1)
                .addData(tableData) // 必选，插入的数据
                .execute(); // 执行
    }
}
