package com.geese.plugin.excel.test;

import com.geese.plugin.excel.StandardWriter;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.util.*;

/**
 * Created by Administrator on 2016/11/12.
 */
public class StandardWriterTest {
    static OutputStream output;

    @BeforeClass
    public static void beforeClass() throws IOException {
        URL url = Thread.currentThread().getContextClassLoader().getResource("demo-writer.xlsx");
        output = new FileOutputStream(url.getFile());
    }

    @Test
    public void test001() {
        List tableData = new ArrayList();
        for (int i = 0; i < 50; i++) {
            Map rowData = new HashMap();
            rowData.put("name", "zhangsan");
            rowData.put("age", i + 10);
            tableData.add(rowData);
        }
        StandardWriter.build(output).insert("0 name, 1 age into 0 limit 3").addData("0", 0, tableData).execute();
    }

    @Test
    public void test002() {
        Map pointData = new HashMap();
        pointData.put("name", "张三");
        pointData.put("age", 18);
        StandardWriter.build(output).insert("{0-1 name, 0-2 age into 0}").addData("0", pointData).execute();
    }

    @Test
    public void test003() {
        List tableData = new ArrayList();
        for (int i = 0; i < 50000; i++) {
            Map rowData = new HashMap();
            rowData.put("name", "zhangsan" + i);
            rowData.put("age", i + 10);
            tableData.add(rowData);
        }
        Map pointData = new HashMap();
        pointData.put("name", "张三");
        pointData.put("age", 18);

        long start = System.currentTimeMillis();
        StandardWriter.build(output)
                .insert("0 name, 1 age into 0 limit 3", "{0-10 name, 0-12 age into 0}")
                .addData("0", 0, tableData)
                .addData("0", pointData)
                .execute();
        long end = System.currentTimeMillis();
        System.out.println(end - start);
    }

    @Test
    public void test004() {
        List tableData1 = new ArrayList();
        for (int i = 0; i < 50; i++) {
            Map rowData = new HashMap();
            rowData.put("name", "zhangsan" + i);
            rowData.put("age", i + 10);
            tableData1.add(rowData);
        }

        List tableData2 = new ArrayList();
        for (int i = 0; i < 50; i++) {
            Map rowData = new HashMap();
            rowData.put("color", "zhangsan" + i);
            rowData.put("birth", new Date());
            tableData2.add(rowData);
        }

        Map pointData = new HashMap();
        pointData.put("name", "张三");
        pointData.put("age", 18);

        long start = System.currentTimeMillis();
        StandardWriter.build(output)
                .insert("0 name, 1 age into 0 limit 3", "5 color, 6 birth into 0 limit 0", "{0-10 name, 0-12 age into 0}")
                .addData("0", 0, tableData1)
                .addData("0", 1, tableData2)
                .addData("0", pointData)
                .execute();
        long end = System.currentTimeMillis();
        System.out.println(end - start);
    }

}
