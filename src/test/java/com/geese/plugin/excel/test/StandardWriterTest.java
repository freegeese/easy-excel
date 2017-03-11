package com.geese.plugin.excel.test;

import com.geese.plugin.excel.StandardWriter;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * <p> 标准Excel写入接口测试 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 21:41
 * @sine 0.0.1
 */
public class StandardWriterTest {
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
     * 实例1：快速上手
     */
    @Test
    public void test001() {
        // 准备表格数据
        List tableData = new ArrayList();
        Map rowData;
        for (int i = 0; i < 50; i++) {
            rowData = new HashMap();
            rowData.put("name", "隔壁老王" + i);
            rowData.put("age", Double.valueOf((Math.random() * 100)).intValue());
            rowData.put("idCard", Double.valueOf((Math.random() * 1000000000)).longValue());
            rowData.put("qq", Double.valueOf((Math.random() * 100000000)).longValue());
            rowData.put("email", Double.valueOf((Math.random() * 1000000)).longValue() + "@qq.com");
            rowData.put("phone", Double.valueOf((Math.random() * 1000000000)).longValue());
            tableData.add(rowData);
        }
        // 把数据插入到excel对应的位置
        StandardWriter.build(output)
                .insert("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone into Sheet1")
                .addData("Sheet1", 0, tableData)
                .execute();
    }


    /**
     * 实例3：列表 + 散列点
     */
    @Test
    public void test003() {
        // 准备表格数据
        List tableData = new ArrayList();
        Map rowData;
        for (int i = 0; i < 50; i++) {
            rowData = new HashMap();
            rowData.put("name", "隔壁老王" + i);
            rowData.put("age", Double.valueOf((Math.random() * 100)).intValue());
            rowData.put("idCard", Double.valueOf((Math.random() * 1000000000)).longValue());
            rowData.put("qq", Double.valueOf((Math.random() * 100000000)).longValue());
            rowData.put("email", Double.valueOf((Math.random() * 1000000)).longValue() + "@qq.com");
            rowData.put("phone", Double.valueOf((Math.random() * 1000000000)).longValue());
            tableData.add(rowData);
        }
        // 准备散列点数据 EASY-EXCEL
        Map pointData = new HashMap();
        pointData.put("e", "E");
        pointData.put("a", "A");
        pointData.put("s", "S");
        pointData.put("y", "Y");

        pointData.put("x", "X");
        pointData.put("c", "C");
        pointData.put("l", "L");

        // 把数据插入到excel对应的位置
        StandardWriter.build(output)
                .insert(
                        "0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone into Sheet1",
                        "{0-8 e, 0-9 a, 0-10 s, 0-11 y, 1-8 e, 1-9 x, 1-10 c, 1-11 e, 1-12 l into Sheet1}"
                )
                .addData("Sheet1", 0, tableData)
                .addData("Sheet1", pointData)
                .execute();
    }

    /**
     * 实例4：多个表格 + 散列点
     */
    @Test
    public void test004() {
        // 准备表格数据
        List tableData1 = new ArrayList();
        Map rowData;
        for (int i = 0; i < 50; i++) {
            rowData = new HashMap();
            rowData.put("name", "隔壁老王" + i);
            rowData.put("age", Double.valueOf((Math.random() * 100)).intValue());
            rowData.put("idCard", Double.valueOf((Math.random() * 1000000000)).longValue());
            rowData.put("qq", Double.valueOf((Math.random() * 100000000)).longValue());
            rowData.put("email", Double.valueOf((Math.random() * 1000000)).longValue() + "@qq.com");
            rowData.put("phone", Double.valueOf((Math.random() * 1000000000)).longValue());
            tableData1.add(rowData);
        }

        List tableData2 = new ArrayList();
        for (int i = 0; i < 20; i++) {
            rowData = new HashMap();
            rowData.put("name1", "隔壁老王" + i);
            rowData.put("age1", Double.valueOf((Math.random() * 100)).intValue());
            rowData.put("idCard1", Double.valueOf((Math.random() * 1000000000)).longValue());
            rowData.put("qq1", Double.valueOf((Math.random() * 100000000)).longValue());
            rowData.put("email1", Double.valueOf((Math.random() * 1000000)).longValue() + "@qq.com");
            rowData.put("phone1", Double.valueOf((Math.random() * 1000000000)).longValue());
            tableData2.add(rowData);
        }

        // 准备散列点数据 EASY-EXCEL
        Map pointData = new HashMap();
        pointData.put("e", "E");
        pointData.put("a", "A");
        pointData.put("s", "S");
        pointData.put("y", "Y");

        pointData.put("x", "X");
        pointData.put("c", "C");
        pointData.put("l", "L");

        // 把数据插入到excel对应的位置
        StandardWriter.build(output)
                .insert(
                        "0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone into Sheet1",
                        "7 name1, 8 age1, 9 idCard1, 10 qq1, 11 email1, 12 phone1 into Sheet1 limit 5",
                        "{0-8 e, 0-9 a, 0-10 s, 0-11 y, 1-8 e, 1-9 x, 1-10 c, 1-11 e, 1-12 l into Sheet1}"
                )
                .addData("Sheet1", 0, tableData1)
                .addData("Sheet1", 1, tableData2)
                .addData("Sheet1", pointData)
                .execute();
    }

    /**
     * 实例5：多sheet插入
     */
    @Test
    public void test005() {
        // 准备表格数据
        List tableData1 = new ArrayList();
        Map rowData;
        for (int i = 0; i < 50; i++) {
            rowData = new HashMap();
            rowData.put("name", "隔壁老王" + i);
            rowData.put("age", Double.valueOf((Math.random() * 100)).intValue());
            rowData.put("idCard", Double.valueOf((Math.random() * 1000000000)).longValue());
            rowData.put("qq", Double.valueOf((Math.random() * 100000000)).longValue());
            rowData.put("email", Double.valueOf((Math.random() * 1000000)).longValue() + "@qq.com");
            rowData.put("phone", Double.valueOf((Math.random() * 1000000000)).longValue());
            tableData1.add(rowData);
        }

        List tableData2 = new ArrayList();
        for (int i = 0; i < 20; i++) {
            rowData = new HashMap();
            rowData.put("name1", "隔壁老王" + i);
            rowData.put("age1", Double.valueOf((Math.random() * 100)).intValue());
            rowData.put("idCard1", Double.valueOf((Math.random() * 1000000000)).longValue());
            rowData.put("qq1", Double.valueOf((Math.random() * 100000000)).longValue());
            rowData.put("email1", Double.valueOf((Math.random() * 1000000)).longValue() + "@qq.com");
            rowData.put("phone1", Double.valueOf((Math.random() * 1000000000)).longValue());
            tableData2.add(rowData);
        }

        // 准备散列点数据 EASY-EXCEL
        Map pointData = new HashMap();
        pointData.put("e", "E");
        pointData.put("a", "A");
        pointData.put("s", "S");
        pointData.put("y", "Y");

        pointData.put("x", "X");
        pointData.put("c", "C");
        pointData.put("l", "L");

        // 把数据插入到excel对应的位置
        StandardWriter.build(output)
                .insert(
                        "0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone into 0",
                        "7 name1, 8 age1, 9 idCard1, 10 qq1, 11 email1, 12 phone1 into 1 limit 5",
                        "{0-8 e, 0-9 a, 0-10 s, 0-11 y, 1-8 e, 1-9 x, 1-10 c, 1-11 e, 1-12 l into 0}",
                        "{0-8 e, 0-9 a, 0-10 s, 0-11 y, 1-8 e, 1-9 x, 1-10 c, 1-11 e, 1-12 l into 1}"
                )
                .addData("0", 0, tableData1)
                .addData("1", 0, tableData2)
                .addData("0", pointData)
                .addData("1", pointData)
                .execute();
    }
}
