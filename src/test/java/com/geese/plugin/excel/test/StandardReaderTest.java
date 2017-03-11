package com.geese.plugin.excel.test;

import com.geese.plugin.excel.StandardWriter;
import org.junit.BeforeClass;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
        // ExcelMapping 表头： 姓名	| 年龄 | 身份证号	| QQ | 邮箱 | 手机
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

}
