package com.geese.plugin.excel.test;

import com.geese.plugin.excel.ExcelSimpleReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelMappingSimpleReaderTest {

    private static InputStream excelFrom;

    @BeforeClass
    public static void beforeClass() throws IOException {
        URL url = Thread.currentThread().getContextClassLoader().getResource("demo-mocai-reader.xlsx");
        excelFrom = new FileInputStream(url.getFile());
    }

    @AfterClass
    public static void afterClass() throws IOException {
        if (null != excelFrom) {
            excelFrom.close();
        }
    }

    @Test
    public void test() throws IOException, InvalidFormatException {
        Object value = ExcelSimpleReader.newInstance(excelFrom)
                .select("7   thick, 8 width, 9 length, 10 weight, 11 unitWeight")
                .from(0, "新")
                .limit(2)
                .execute();
        System.out.println(value);

    }


}
