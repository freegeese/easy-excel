# 像SQL一样操作Excel
- 简化对Excel的读写操作
- 尽可能的实现能够像SQL一样操作来操作Excel

##快速上手

###对Excel的简单read，使用SimpleReader类完成
```
package com.geese.plugin.excel.test;

import com.geese.plugin.excel.filter.CellAfterReadFilter;
import com.geese.plugin.excel.filter.RowAfterReadFilter;
import com.geese.plugin.excel.SimpleReader;
import com.geese.plugin.excel.SimpleWriter;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.filter.CellBeforeReadFilter;
import com.geese.plugin.excel.filter.RowBeforeReadFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.*;

/**
 * SimpleReader 接口测试
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 18:05
 * @sine 0.0.1
 */
public class SimpleReaderTest {

    private static InputStream input;

    @BeforeClass
    public static void beforeClass() throws IOException {
        URL url = Thread.currentThread().getContextClassLoader().getResource("demo-reader.xlsx");
        // 准备数据
        // Excel 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
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
        // 构建一个输出流，向被读取的excel写入测试数据
        FileOutputStream output = new FileOutputStream(url.getFile());
        SimpleWriter.build(output)
                .insert("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")
                .limit(1)
                .addData(tableData)
                .execute();
        output.flush();
        output.close();

        // 构建一个输入流，读取excel数据
        input = new FileInputStream(url.getFile());
    }

    @AfterClass
    public static void afterClass() throws IOException {
        if (null != input) {
            input.close();
        }
    }

    /**
     * 实例1：快速使用
     */
    @Test
    public void test001() {
        Collection result = SimpleReader.build(input)
                .select("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")
                .execute();
        System.out.println(result);
    }

    /**
     * 实力2：可选配置
     */
    @Test
    public void test002() {
        Collection result = SimpleReader.build(input)
                .select("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")
                .from("0")      // 可选（默认是从第0个sheet读取, [如果是数字会优先使用名称获取sheet，没找到才会使用下标获取sheet]）
                .limit(3, 5)    // 可选（默认从定义0行开始读取，读取所有行）
                .execute();
        System.out.println(result.size());
    }

    /**
     * 实例3：where条件帅选，使用占位符参数
     */
    @Test
    public void test003() {
        // 占位符参数
        Collection result = SimpleReader.build(input)
                .select("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")
                .where("name like ? and (age > ? and qq like ?) or name in ?")
                .addParameter(Arrays.asList("%鲁%", 50, "5%", Arrays.asList("张白萱", "张若云", "张雅畅", "张雅寒", "张雨华")))
                .execute();
        System.out.println(result.size());
    }

    /**
     * 实例4：where条件帅选，使用命名的参数
     */
    @Test
    public void test004() {
        // 命名的参数
        Map namedParameter = new HashMap();
        namedParameter.put("name", "%鲁%");
        namedParameter.put("age", 26);
        namedParameter.put("qq", "5%");
        namedParameter.put("names", Arrays.asList("张白萱", "张若云", "张雅畅", "张雅寒", "张雨华"));
        Collection result = SimpleReader.build(input)
                .select("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")
                .where("name like :name and (age > :age and qq like :qq) or name in :names")
                .addParameter(namedParameter)
                .execute();
        System.out.println(result.size());
    }

    /**
     * 实例5：过滤器
     */
    @Test
    public void test005() {
        Collection result = SimpleReader.build(input)
                .select("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")
                .addFilter(new RowBeforeReadFilter() {
                    // 可以对 row 进行修改
                    @Override
                    public void doFilter(Row target, Object data, Table config) {
                        System.out.println("<<<<<<<<<<<<<<<<读取行之前过滤：" + data);
                    }
                }, new RowAfterReadFilter() {
                    // 可以对 data 进行修改
                    @Override
                    public void doFilter(Row target, Object data, Table config) {
                        System.out.println("读取行之后过滤：" + data + ">>>>>>>>>>>>>>>>>");
                    }
                }, new CellBeforeReadFilter() {
                    // 可以对 cell 进行修改
                    @Override
                    public void doFilter(Cell target, Object data, Point config) {
                        System.out.println("<<<<<<<<<<<<<<<<读单元格之前过滤：" + data);
                    }
                }, new CellAfterReadFilter() {
                    // 可以对 data 进行修改
                    @Override
                    public void doFilter(Cell target, Object data, Point config) {
                        System.out.println("读单元格之后过滤：" + data + ">>>>>>>>>>>>>>>>>");
                    }
                })
                .execute();
        System.out.println(result);
    }
}

```
### 对Excel的简单write，使用SimpleWriter类完成
```
package com.geese.plugin.excel.test;

import com.geese.plugin.excel.SimpleWriter;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.filter.CellWriteFilter;
import com.geese.plugin.excel.filter.RowWriteFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
        // Excel 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
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
        // Excel 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
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
     * 实例3：使用单元格过滤器和行过滤器
     */
    @Test
    public void test003() {
        // Excel 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
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
        /**
         * 写入到row之前过滤，可对row和data进行修改
         */
        RowWriteFilter rowWriteFilter = new RowWriteFilter() {
            @Override
            public void doFilter(Row target, Object data, Table config) {
                System.out.println(data);
            }
        };
        /**
         * 写入到cell之前过滤，可对cell和data进行修改
         */
        CellWriteFilter cellWriteFilter = new CellWriteFilter() {
            @Override
            public void doFilter(Cell target, Object data, Point config) {
                System.out.println(data);
            }
        };
        // 通过SimpleWriter类操作
        SimpleWriter.build(output)  // 必选，将生成的excel输出到什么地方
                .insert("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")  // 必选，数据Key与Excel列的映射
                .addFilter(rowWriteFilter, cellWriteFilter)
                .addData(tableData) // 必选，插入的数据
                .execute(); // 执行
    }

    /**
     * 实例4：使用模板
     */
    @Test
    public void test004() {
        // Excel 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
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

```
## 接口说明
- 暂无

## 高级部分
### StandardReader 接口使用
```
package com.geese.plugin.excel.test;

import com.geese.plugin.excel.StandardReader;
import com.geese.plugin.excel.StandardWriter;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.filter.CellAfterReadFilter;
import com.geese.plugin.excel.filter.CellBeforeReadFilter;
import com.geese.plugin.excel.filter.RowAfterReadFilter;
import com.geese.plugin.excel.filter.RowBeforeReadFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.*;

/**
 * <p> 标准Excel读取接口测试 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 22:36
 * @sine 0.0.1
 */
public class StandardReaderTest {
    static InputStream input;

    @BeforeClass
    public static void beforeClass() throws IOException {
        URL url = Thread.currentThread().getContextClassLoader().getResource("demo-reader.xlsx");
        // 准备数据
        // Excel 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
        // 准备数据, 每一行是一个Map, 每一个表格是一个List<Map>
        String[] names = "鲁沛儿 鲁天薇 鲁飞雨 鲁天纵 鲁白梦 鲁嘉胜 鲁盼巧 鲁访天 鲁清妍 鲁盼晴 张馨蓉 张白萱 张若云 张雅畅 张雅寒 张雨华".split("\\s+");
        List<Map> tableData = new ArrayList<>();
        Map rowData;
        for (int i = 0; i < 100; i++) {
            rowData = new HashMap();
            rowData.put("name", names[Double.valueOf(Math.random() * names.length).intValue()]);
            rowData.put("age", Double.valueOf(Math.random() * 100).intValue());
            rowData.put("idCard", Double.valueOf(Math.random() * 1000000000).longValue());
            rowData.put("qq", Double.valueOf(Math.random() * 1000000000).longValue());
            rowData.put("email", Double.valueOf(Math.random() * 1000000000).longValue() + "@163.com");
            rowData.put("phone", Double.valueOf(Math.random() * 1000000000).longValue());
            tableData.add(rowData);
        }
        // 构建一个输出流，向被读取的excel写入测试数据
        FileOutputStream output = new FileOutputStream(url.getFile());
        StandardWriter.build(output)
                .insert(
                        "0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone into 0",
                        "0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone into 1"
                )
                .addData("0", 0, tableData)
                .addData("1", 0, tableData)
                .execute();
        output.flush();
        output.close();

        // 构建一个输入流，读取excel数据
        input = new FileInputStream(url.getFile());
    }

    /**
     * 实例1：快速上手
     */
    @Test
    public void test001() {
        Object result = StandardReader
                .build(input)
                .select("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone from 0")
                .execute();
        System.out.println(result);
    }

    /**
     * 实例2：可选配置
     */
    @Test
    public void test002() {
        StandardReader
                .build(input)
                // limit: [startRow, size] 从哪行开始读，读取多少行
                .select("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone from 0 limit 0, 10")
                // 绑定过滤器到一个sheet上
                .addFilter("0", new RowBeforeReadFilter() {
                    @Override
                    public void doFilter(Row target, Object data, Table config) {
                        System.out.println("<><><><><><><>读取Row之前过滤：" + data + "<><><><><><><>");
                    }
                }, new RowAfterReadFilter() {
                    @Override
                    public void doFilter(Row target, Object data, Table config) {
                        System.out.println("<><><><><><><>读取Row之后过滤：" + data + "<><><><><><><>");
                    }
                }, new CellBeforeReadFilter() {
                    @Override
                    public void doFilter(Cell target, Object data, Point config) {
                        System.out.println("<><><><><><><>读取Cell之前过滤：" + data + "<><><><><><><>");
                    }
                }, new CellAfterReadFilter() {
                    @Override
                    public void doFilter(Cell target, Object data, Point config) {
                        System.out.println("<><><><><><><>读取Cell之后过滤：" + data + "<><><><><><><>");
                    }
                })
                .execute();
    }

    /**
     * 实例3：where 条件过滤
     */
    @Test
    public void test003() {
        Map namedParameter = new HashMap();
        namedParameter.put("name", "鲁%");
        namedParameter.put("age", 20);
        namedParameter.put("qq", "%12%");
        namedParameter.put("names", Arrays.asList("张馨蓉", "张白萱", "张若云"));
        Object result = StandardReader
                .build(input)
                // where 条件过滤，支持占位符和命名参数
                .select(
                        "0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone from 0 where name like ? and (age > ? or qq like ? or name in ?)",
                        "0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone from 1 where name like :name and (age > :age or qq like :qq or name in :names)"
                )
                // 添加占位符参数
                .addParameter("0", 0, Arrays.asList("鲁%", 20, "%12%", Arrays.asList("张馨蓉", "张白萱", "张若云")))
                // 添加命名的参数
                .addParameter("1", 0, namedParameter)
                .execute();
        System.out.println(result);
    }

    /**
     * 实例4：散列点
     */
    @Test
    public void test004() {
        Object result = StandardReader
                .build(input)
                .select(
                        "{0-1 name, 0-2 age from 0}",
                        "{1-1 name, 1-2 age from 1}"
                )
                .execute();
        System.out.println(result);
    }
}
```

### StandardWriter 接口使用
```
package com.geese.plugin.excel.test;

import com.geese.plugin.excel.StandardWriter;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.filter.CellWriteFilter;
import com.geese.plugin.excel.filter.RowWriteFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
     * 实例11：可选配置项
     */
    @Test
    public void test002() {
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
        StandardWriter
                // template: 使用模板来接收写入的数据
                .build(output, template)
                // limit:[startRow, size] 从哪行开始写，写多少行 默认：[0, tableData.size()]
                .insert("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone into Sheet1 limit 1, 30")
                // filter: 在数据写入到row或cell之前，可以对数据进行过滤修改, 过滤器需要绑定到某个Sheet上执行
                .addFilter("Sheet1", new RowWriteFilter() {
                    @Override
                    public void doFilter(Row target, Object data, Table config) {
                        System.out.println("<<<<<<<<<<<<写入Row之前过滤：" + data + ">>>>>>>>>>>");
                    }
                }, new CellWriteFilter() {
                    @Override
                    public void doFilter(Cell target, Object data, Point config) {
                        System.out.println("<<<<<<<<<<<<写入Cell之前过滤：" + data + ">>>>>>>>>>>");
                    }
                })
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
```
