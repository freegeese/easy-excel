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
