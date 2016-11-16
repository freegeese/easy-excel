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
