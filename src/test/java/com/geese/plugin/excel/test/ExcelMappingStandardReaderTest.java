package com.geese.plugin.excel.test;

import com.geese.plugin.excel.ExcelStandardReader;
import com.geese.plugin.excel.ExcelTemplate;
import com.geese.plugin.excel.filter.read.RowAfterReadFilter;
import com.geese.plugin.excel.filter.read.RowBeforeReadFilter;
import com.geese.plugin.excel.filter.read.SheetBeforeReadFilter;
import com.geese.plugin.excel.mapping.SheetMapping;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Collection;
import java.util.Map;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelMappingStandardReaderTest {

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
    public void test1() throws IOException, InvalidFormatException {
        Object result = ExcelStandardReader.newInstance(excelFrom)
                .select("7 thick, 8 width, 9 length, 10 weight, 11 unitWeight from Sheet1 limit 10,10")
                .execute();
        System.out.println(result);
    }

    @Test
    public void test2() throws IOException, InvalidFormatException {
        Object result = ExcelStandardReader.newInstance(excelFrom)
                .select(
                        "7 thick, 8 width, 9 length, 10 weight, 11 unitWeight from Sheet1 limit 10,10",
                        "{0-0 no, 0-1 type from Sheet2}"
                )
                .execute();
        System.out.println(result);
    }

    @Test
    public void test3() throws IOException, InvalidFormatException {
        Object result = ExcelStandardReader.newInstance(excelFrom)
                .select(
                        "7 thick, 8 width, 9 length, 10 weight, 11 unitWeight from Sheet1 limit 10,10",
                        "{0-0 no, 0-1 type from Sheet2}"
                )
                .filter(new RowBeforeReadFilter() {
                    @Override
                    public boolean doFilter(Row target, Object data, SheetMapping mapping) {
                        return target.getRowNum() <= 15;
                    }
                }, "Sheet1")
                .filter(new RowAfterReadFilter() {
                    @Override
                    public boolean doFilter(Row target, Object data, SheetMapping mapping) {
                        System.out.println(data);
                        return target.getRowNum() <= 14;
                    }
                }, "Sheet1")
                .filter(new SheetBeforeReadFilter() {
                    @Override
                    public boolean doFilter(Sheet target, Object data, SheetMapping mapping) {
                        System.out.println(target.getSheetName());
                        return true;
                    }
                }, "Sheet1")
                .execute();
        System.out.println(result);
    }

    @Test
    public void test4() throws IOException, InvalidFormatException {
        Map result = ExcelStandardReader.newInstance(excelFrom)
                .select(
                        "no, type, shape, gy, jonggong from Sheet1 limit 10,100",
                        "{0-0 no, 0-1 type from Sheet2}"
                )
                .execute();
        Map sheet1 = (Map) result.get("Sheet1");
        Collection o1 = (Collection) sheet1.get(ExcelTemplate.TABLE_DATA_KEY);
        for (Object o : o1) {
            System.out.println(o);
        }
        System.out.println(o1);

        Map sheet2 = (Map) result.get("Sheet2");
        Object o = sheet2.get(ExcelTemplate.POINT_DATA_KEY);
        System.out.println(o);
    }

}
