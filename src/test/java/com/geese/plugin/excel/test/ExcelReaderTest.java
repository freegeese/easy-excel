package com.geese.plugin.excel.test;

import com.geese.plugin.excel.ExcelReader;
import com.geese.plugin.excel.ExcelResult;
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
public class ExcelReaderTest {

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
        Object result = ExcelReader.newInstance(excelFrom)
                .select("7 thick, 8 width, 9 length, 10 weight, 11 unitWeight from Sheet1 limit 10,10")
                .execute();
        System.out.println(result);
    }

    @Test
    public void test2() throws IOException, InvalidFormatException {
        Object result = ExcelReader.newInstance(excelFrom)
                .select(
                        "7 thick, 8 width, 9 length, 10 weight, 11 unitWeight from Sheet1 limit 10,10",
                        "{0-0 no, 0-1 type from Sheet2}"
                )
                .execute();
        System.out.println(result);
    }

    @Test
    public void test3() throws IOException, InvalidFormatException {
        Object result = ExcelReader.newInstance(excelFrom)
                .select(
                        "7 thick, 8 width, 9 length, 10 weight, 11 unitWeight from Sheet1 limit 10,10",
                        "{0-0 no, 0-1 type from Sheet2}"
                )
                .filter(new RowBeforeReadFilter() {
                    @Override
                    public boolean doFilter(Row target, Object data, SheetMapping mapping, Map context) {
                        return target.getRowNum() <= 15;
                    }
                }, "Sheet1")
                .filter(new RowAfterReadFilter() {
                    @Override
                    public boolean doFilter(Row target, Object data, SheetMapping mapping, Map context) {
                        System.out.println(data);
                        return target.getRowNum() <= 14;
                    }
                }, "Sheet1")
                .filter(new SheetBeforeReadFilter() {
                    @Override
                    public boolean doFilter(Sheet target, Object data, SheetMapping mapping, Map context) {
                        System.out.println(target.getSheetName());
                        return true;
                    }
                }, "Sheet1")
                .execute();
        System.out.println(result);
    }

    @Test
    public void test4() throws IOException, InvalidFormatException {
        ExcelResult result = ExcelReader.newInstance(excelFrom)
                .select(
                        "no, type, shape, gy, jonggong from Sheet1 limit 10,100",
                        "{0-0 no, 0-1 type from Sheet2}"
                )
                .filter(new RowAfterReadFilter() {
                    @Override
                    public boolean doFilter(Row target, Object data, SheetMapping mapping, Map context) {
                        if (target.getRowNum() < 50) {
                            context.put(target.getRowNum(), data);
                        }
                        return true;
                    }
                }, "Sheet1")
                .execute();
        Collection sheet1TableData = result.getTableData("Sheet1");
        for (Object o : sheet1TableData) {
            System.out.println(o);
        }
        System.out.println(sheet1TableData);

        Map sheet2PointData = result.getPointData("Sheet2");
        System.out.println(sheet2PointData);

        System.out.println(result.getContext());
    }

}
