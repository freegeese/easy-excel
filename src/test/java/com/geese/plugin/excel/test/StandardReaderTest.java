package com.geese.plugin.excel.test;

import com.geese.plugin.excel.StandardReader;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Arrays;

/**
 * Created by Administrator on 2016/11/12.
 */
public class StandardReaderTest {
    static InputStream input;

    @BeforeClass
    public static void beforeClass() throws IOException {
        URL url = Thread.currentThread().getContextClassLoader().getResource("demo-reader.xlsx");
        input = new FileInputStream(url.getFile());
    }

    @Test
    public void test001() {
        StandardReader.build(input).select("0 name,1 age from 0 where age >= ? and name like ? limit 0").addParameter("0", 0, Arrays.asList(18, "张")).execute();
    }

    @Test
    public void test002() {
        StandardReader.build(input).select("0 name,1 age from 0", "{0-0 name, 0-1 age from 0}").execute();
    }

}
