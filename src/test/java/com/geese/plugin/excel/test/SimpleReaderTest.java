package com.geese.plugin.excel.test;

import com.geese.plugin.excel.config.MySheet;
import com.geese.plugin.excel.filter.*;
import com.geese.plugin.excel.SimpleReader;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.Table;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileInputStream;
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
        URL url = Thread.currentThread().getContextClassLoader().getResource("demo-mocai-reader.xlsx");
//        // 准备数据
//        // Excel 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
//        // 准备数据, 每一行是一个Map, 每一个表格是一个List<Map>
//        String names = "鲁沛儿 鲁天薇 鲁飞雨 鲁天纵 鲁白梦 鲁嘉胜 鲁盼巧 鲁访天 鲁清妍 鲁盼晴 张馨蓉 张白萱 张若云 张雅畅 张雅寒 张雨华";
//        List<Map> tableData = new ArrayList<>();
//        Map rowData;
//        for (String name : names.split("\\s+")) {
//            rowData = new HashMap();
//            rowData.put("name", name);
//            rowData.put("age", Double.valueOf(Math.random() * 100).intValue());
//            rowData.put("idCard", Double.valueOf(Math.random() * 1000000000).longValue());
//            rowData.put("qq", Double.valueOf(Math.random() * 1000000000).longValue());
//            rowData.put("email", Double.valueOf(Math.random() * 1000000000).longValue() + "@163.com");
//            rowData.put("phone", Double.valueOf(Math.random() * 1000000000).longValue());
//            tableData.add(rowData);
//        }
//        // 构建一个输出流，向被读取的excel写入测试数据
//        FileOutputStream output = new FileOutputStream(url.getFile());
//        SimpleWriter.build(output)
//                .insert("0 name, 1 age, 2 idCard, 3 qq, 4 email, 5 phone")
//                .limit(1)
//                .addData(tableData)
//                .execute();
//        output.flush();
//        output.close();

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
                .limit(3, 1)    // 可选（默认从定义0行开始读取，读取所有行）
                .execute();
        System.out.println(result);
    }

    @Test
    public void test003() {
        SimpleReader.build(input)
                .select("7 thick, 8 width, 9 length, 10 weight, 11 unitWeight")
                .limit(3, 2)
                .addFilter(new CellBeforeReadFilter() {
                    @Override
                    public boolean doFilter(Cell target, Object data, Point config) {
                        System.out.println("+++++++++++" + target.getRowIndex() + ":" + target.getColumnIndex() + "++++++++++");
                        return true;
                    }
                })
                .addFilter(new CellAfterReadFilter() {
                    @Override
                    public boolean doFilter(Cell target, Object data, Point config) {
                        System.out.println(data);
                        return true;
                    }
                })
                .addFilter(new RowBeforeReadFilter() {
                    @Override
                    public boolean doFilter(Row target, Object data, Table config) {
                        System.out.println("+++++++++++++" + target.getRowNum() + "+++++++++++++");
                        return true;
                    }
                })
                .addFilter(new RowAfterReadFilter() {
                    @Override
                    public boolean doFilter(Row target, Object data, Table config) {
                        System.out.println(data);
                        return true;
                    }
                })
                .addFilter(new SheetBeforeReadFilter() {
                    @Override
                    public boolean doFilter(Sheet target, Object data, MySheet config) {
                        System.out.println("++++++++" + target.getSheetName() + "++++++++");
                        return true;
                    }
                })
                .addFilter(new SheetAfterReadFilter() {
                    @Override
                    public boolean doFilter(Sheet target, Object data, MySheet config) {
                        System.out.println(data);
                        return false;
                    }
                })
                .execute();
    }


}
