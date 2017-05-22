package com.geese.plugin.excel.test;

import com.geese.plugin.excel.ExcelResult;
import com.geese.plugin.excel.ExcelValidation;
import com.geese.plugin.excel.ExcelWriter;
import com.geese.plugin.excel.filter.WriteFilter;
import com.geese.plugin.excel.filter.write.RowAfterWriteFilter;
import com.geese.plugin.excel.filter.write.RowBeforeWriteFilter;
import com.geese.plugin.excel.mapping.SheetMapping;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelWriterTest {

    private static OutputStream excelOutput;
    private static File excelTemplate;

    @BeforeClass
    public static void beforeClass() throws IOException {
        URL url = Thread.currentThread().getContextClassLoader().getResource("demo-writer.xlsx");
        URL templateUrl = Thread.currentThread().getContextClassLoader().getResource("demo-writer-template.xlsx");
        excelOutput = new FileOutputStream(url.getFile());
        excelTemplate = new File(templateUrl.getFile());
    }

    @AfterClass
    public static void afterClass() throws IOException {
        if (null != excelOutput) {
            excelOutput.close();
        }
    }

    @Test
    public void testSimple() throws IOException, InvalidFormatException {
        ExcelWriter.newInstance(excelOutput)
                .insert("string, double, float, long, integer, boolean, date into Sheet1")
                .addData(generateTableData(), "Sheet1")
                .useXlsxFormat()
                .execute();
    }

    /**
     * 测试limit: first parameter: start row number, second parameter: row interval
     *
     * @throws IOException
     * @throws InvalidFormatException
     */
    @Test
    public void testLimit() throws IOException, InvalidFormatException {
        ExcelWriter.newInstance(excelOutput)
                .insert("string, double, float, long, integer, boolean, date into Sheet1 limit 10,1")
                .addData(generateTableData(), "Sheet1")
                .useXlsxFormat()
                .execute();
    }

    /**
     * 指定列位置
     *
     * @throws IOException
     * @throws InvalidFormatException
     */
    @Test
    public void testSpecifyColumnNumber() throws IOException, InvalidFormatException {
        ExcelWriter.newInstance(excelOutput)
                .insert("2 string, double, 5 float, long, integer, boolean, date into Sheet1 limit 10,1")
                .addData(generateTableData(), "Sheet1")
                .useXlsxFormat()
                .execute();
    }

    @Test
    public void testFilter() throws IOException, InvalidFormatException {
        ExcelResult result = ExcelWriter.newInstance(excelOutput)
                .insert("2 string, double, 5 float, long, integer, boolean, date into Sheet1 limit 10,1")
                .addData(generateTableData(), "Sheet1")
                .filters(new WriteFilter[]{
                        new RowBeforeWriteFilter() {
                            @Override
                            public boolean doFilter(Row target, Object data, SheetMapping mapping, Map context) {
                                context.put("RowBeforeWriteFilter--> " + target.getRowNum(), data);
                                System.out.println("RowBeforeWriteFilter--> " + target.getRowNum());
                                return true;
                            }
                        },
                        new RowAfterWriteFilter() {
                            @Override
                            public boolean doFilter(Row target, Object data, SheetMapping mapping, Map context) {
                                context.put("RowAfterWriteFilter--> " + target.getRowNum(), data);
                                System.out.println("RowAfterWriteFilter--> " + target.getRowNum());
                                return true;
                            }
                        }
                }, "Sheet1")
                .useXlsxFormat()
                .execute();
        // 一次写操作的上下文信息
        Map context = result.getContext();
        System.out.println(context);
    }

    @Test
    public void testTemplate() throws IOException, InvalidFormatException {
        ExcelResult result = ExcelWriter.newInstance(excelOutput)
                .insert("2 string, double, 5 float, long, integer, boolean, date into Sheet1 limit 1,2")
                .addData(generateTableData(), "Sheet1")
                .setTemplate(excelTemplate)
                .useXlsxFormat()
                .execute();
        // 一次写操作的上下文信息
        Map context = result.getContext();
        System.out.println(context);
    }

    @Test
    public void testTemplateValidation() throws IOException, InvalidFormatException {
        ExcelWriter.newInstance(excelOutput)
                .setTemplate(excelTemplate)
                .addValidation(new ExcelValidation(1, 20, 0, 0, Arrays.asList("1", "2")), "0")
                .execute();
    }

    @Test
    public void testPoint() throws IOException, InvalidFormatException {
        ExcelWriter.newInstance(excelOutput)
                .insert("{1-2 name, 3-4 birthday} into Sheet1}")
                .addData(generatePointData(), "Sheet1")
                .setTemplate(excelTemplate)
                .useXlsxFormat()
                .execute();
    }

    @Test
    public void testPicture() throws IOException, InvalidFormatException {
        ExcelWriter.newInstance(excelOutput)
                .insert("p1, p2, p3 into Sheet1", "{0-6 p6, 0-7 p7 into Sheet1}")
                .addData(generateTablePictureData(), "Sheet1")
                .addData(generatePointPictureData(), "Sheet1")
                .setTemplate(excelTemplate)
                .useXlsxFormat()
                .execute();
    }

    private List<Map> generateTablePictureData() throws IOException {
        URL url = Thread.currentThread().getContextClassLoader().getResource("images");
        String path = url.getPath();
        File imagesDirectory = new File(path);
        File[] files = imagesDirectory.listFiles();
        List<Map> returnValue = new ArrayList<>();
        Map<String, Object> rowData = new LinkedHashMap<>();
        returnValue.add(rowData);
        int index = 1;
        for (File file : files) {
            if (index % 5 == 0) {
                rowData = new LinkedHashMap<>();
                returnValue.add(rowData);
                index = 1;
            }
            byte[] bytes = Files.readAllBytes(Paths.get(file.getPath()));
            rowData.put("p" + (index++), bytes);
        }
        return returnValue;
    }

    private Map generatePointPictureData() throws IOException {
        URL url = Thread.currentThread().getContextClassLoader().getResource("images");
        String path = url.getPath();
        File imagesDirectory = new File(path);
        File[] files = imagesDirectory.listFiles();
        Map<String, Object> pointPictureData = new LinkedHashMap<>();
        int index = 1;
        for (File file : files) {
            byte[] bytes = Files.readAllBytes(Paths.get(file.getPath()));
            pointPictureData.put("p" + (index++), bytes);
        }
        return pointPictureData;
    }

    private Map generatePointData() {
        Map<String, Object> pointData = new LinkedHashMap<>();
        pointData.put("name", "你好中国，My name is zhangguangyong!^^");
        pointData.put("birthday", new Date());
        return pointData;
    }

    private List<Map> generateTableData() {
        List<Map> tableData = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, Object> rowData = new HashMap<>();
            rowData.put("string", "你好中国，My name is zhangguangyong!^^");
            rowData.put("double", Double.valueOf(Math.random() * 1000000000));
            rowData.put("float", Double.valueOf(Math.random() * 1000000000).floatValue());
            rowData.put("long", Double.valueOf(Math.random() * 100000000).longValue());
            rowData.put("integer", Double.valueOf(Math.random() * 100000000).intValue());
            rowData.put("boolean", Math.random() * 10 > 5);
            rowData.put("date", new Date(System.currentTimeMillis() + Double.valueOf(Math.random() * 1000000000).longValue()));
            tableData.add(rowData);
        }
        return tableData;
    }

}
