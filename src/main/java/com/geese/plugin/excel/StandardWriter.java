package com.geese.plugin.excel;

import com.geese.plugin.excel.config.Excel;
import com.geese.plugin.excel.config.MySheet;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.core.ExcelHelper;
import com.geese.plugin.excel.core.ExcelSupport;
import com.geese.plugin.excel.core.OperationKey;
import com.geese.plugin.excel.filter.*;
import com.geese.plugin.excel.util.Check;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * <p>标准Excel写入操作，提供比较全面的写入接口 <br>
 * <ul>
 * <li>
 * 单个sheet单个table写入
 * {@code insert("0 name, 1 age, 2 idCard into Sheet1 limit 1").addData("Sheet1", 0, tableData)}
 * </li>
 * <li>
 * 单个sheet多个table写入
 * {@code insert(
 * "0 name, 1 age, 2 idCard into Sheet1 limit 1"
 * "4 namex, 5 agex, 6 idCardx into Sheet1 limit 1",
 * )
 * .addData("Sheet1", 0, tableData1)}
 * .addData("Sheet1", 1, tableData2)}
 * </li>
 * <li>
 * 多个sheet多个table写入
 * {@code insert(
 * "0 name, 1 age, 2 idCard into Sheet1 limit 1",
 * "4 namex, 5 agex, 6 idCardx into 1 limit 1"
 * )
 * .addData("Sheet1", 0, tableData1)}
 * .addData("1", 0, tableData2)}
 * </li>
 * <li>
 * 多个sheet多个table多个散列点写入
 * {@code insert(
 * "0 name, 1 age, 2 idCard into Sheet1 limit 1",
 * "4 namex, 5 agex, 6 idCardx into 1 limit 1",
 * "{3-4 name2, 0-9 age3 into Sheet1}",
 * "{3-4 name3, 0-9 age4 into 1}",
 * )
 * .addData("Sheet1", 0, tableData1)
 * .addData("1", 0, tableData2)
 * .addData("Sheet1", pointData1)
 * .addData("1", pointData2)
 * }
 * </li>
 * </ul>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 17:35
 * @sine 0.0.1
 */
public class StandardWriter {

    /**
     * 写入输入到Excel后，输出到哪里
     */
    private OutputStream output;

    /**
     * 当不存在模板的时候，使用什么格式创建excel
     */
    private boolean useXlsx = true;

    /**
     * 写入Excel所需的Sheet配置信息
     */
    private Map<String, MySheet> sheetConfigMap;

    /**
     * 写入Excel时使用的模板
     */
    private InputStream template;

    public StandardWriter() {
        this.sheetConfigMap = new LinkedHashMap();
    }

    public static StandardWriter build(OutputStream output) {
        Check.notNull(output);
        return build(output, true);
    }

    public static StandardWriter build(OutputStream output, boolean useXlsx) {
        Check.notNull(output);
        StandardWriter writer = new StandardWriter();
        writer.output = output;
        writer.useXlsx = useXlsx;
        return writer;
    }

    public static StandardWriter build(OutputStream output, File template) {
        Check.notNull(output, template);
        try {
            return build(output, new FileInputStream(template));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new IllegalArgumentException("模板文件不存在：" + template);
        }
    }

    public static StandardWriter build(OutputStream output, InputStream template) {
        Check.notNull(output, template);
        StandardWriter writer = new StandardWriter();
        writer.output = output;
        writer.template = template;
        return writer;
    }

    /**
     * 通过插入语句与构建SheetConfig信息
     *
     * @param firstInsert
     * @param more
     * @return
     */
    public StandardWriter insert(String firstInsert, String... more) {
        Check.notEmpty(firstInsert);
        List<String> inserts = new ArrayList();
        inserts.add(firstInsert);
        if (null != more) {
            inserts.addAll(Arrays.asList(more));
        }
        // 遍历插入语句
        for (String insert : inserts) {
            // 散列点插入语句 insert point
            if (insert.matches("^\\{.+\\}$")) {
                insert = insert.replaceAll("\\{|\\}", "");
                Map<OperationKey, String> keyDataMap = ExcelHelper.insertKeyParse(insert);
                String sheet = keyDataMap.get(OperationKey.INTO);
                MySheet sheat = sheetConfigMap.get(sheet);
                if (null == sheat) {
                    sheat = new MySheet();
                    ExcelHelper.setSheet(sheet, sheat);
                    sheetConfigMap.put(sheet, sheat);
                }
                // [column row name]
                String[] rowColumnNames = keyDataMap.get(OperationKey.COLUMN).split(",");
                for (String rowColumnName : rowColumnNames) {
                    String[] items = rowColumnName.trim().split("-|\\s+");
                    Point point = new Point();
                    point.setX(Integer.valueOf(items[0]));
                    point.setY(Integer.valueOf(items[1]));
                    point.setKey(items[2]);
                    sheat.addPoint(point);
                    point.setMySheet(sheat);
                }
                continue;
            }

            // 列表插入 insert table
            Map<OperationKey, String> keyDataMap = ExcelHelper.insertKeyParse(insert);
            // into sheet
            String sheet = keyDataMap.get(OperationKey.INTO);
            MySheet sheat = sheetConfigMap.get(sheet);
            if (null == sheat) {
                sheat = new MySheet();
                ExcelHelper.setSheet(sheet, sheat);
                sheetConfigMap.put(sheet, sheat);
            }

            // insert columns
            Table table = new Table();
            String[] columnNames = keyDataMap.get(OperationKey.COLUMN).split(",");
            for (String columnName : columnNames) {
                String[] columnWithName = columnName.trim().split("\\s+");
                Point point = new Point();
                point.setY(Integer.valueOf(columnWithName[0]));
                point.setKey(columnWithName[1]);
                table.addColumn(point);
                point.setTable(table);
            }

            // limit
            if (keyDataMap.containsKey(OperationKey.LIMIT)) {
                String[] startWithSize = keyDataMap.get(OperationKey.LIMIT).split(",");
                table.setStartRow(Integer.valueOf(startWithSize[0].trim()));
                if (startWithSize.length > 1) {
                    table.setEndRow(table.getStartRow() + Integer.valueOf(startWithSize[1].trim()));
                }
            }

            sheat.addTable(table);
            table.setMySheet(sheat);
        }
        return this;
    }

    /**
     * 添加Table数据
     *
     * @param sheet
     * @param tableIndex
     * @param tableData
     * @return
     */
    public StandardWriter addData(String sheet, Integer tableIndex, List tableData) {
        Check.notEmpty(sheet, tableIndex, tableData);
        if (!sheetConfigMap.containsKey(sheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + sheet);
        }
        MySheet mySheetConfig = sheetConfigMap.get(sheet);
        Table table = mySheetConfig.getTables().get(tableIndex);
        table.setData(tableData);
        return this;
    }

    /**
     * 添加散列点数据
     *
     * @param sheet
     * @param pointData
     * @return
     */
    public StandardWriter addData(String sheet, Map<String, Object> pointData) {
        Check.notEmpty(sheet, pointData);
        if (!sheetConfigMap.containsKey(sheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + sheet);
        }
        MySheet mySheetConfig = sheetConfigMap.get(sheet);
        Set<String> keys = pointData.keySet();
        for (String key : keys) {
            Point point = mySheetConfig.findPoint(key);
            Check.notNull(point, "找不到：[" + key + "] 对应的point");
            point.setData(pointData.get(key));
        }
        return this;
    }

    public StandardWriter addFilter(String sheet, Integer tableIndex, Collection<Filter> filters) {
        Check.notEmpty(filters);
        if (!sheetConfigMap.containsKey(sheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + sheet);
        }
        MySheet mySheetConfig = sheetConfigMap.get(sheet);
        Table table = mySheetConfig.getTables().get(tableIndex);

        for (Filter filter : filters) {
            if (filter instanceof RowBeforeWriteFilter) {
                table.addRowBeforeWriteFilter((RowBeforeWriteFilter) filter);
                continue;
            }
            if (filter instanceof RowAfterWriteFilter) {
                table.addRowAfterWriteFilter((RowAfterWriteFilter) filter);
                continue;
            }
            if (filter instanceof CellBeforeWriteFilter) {
                table.addCellBeforeWriteFilter((CellBeforeWriteFilter) filter);
                continue;
            }
            if (filter instanceof CellAfterWriteFilter) {
                table.addCellAfterWriteFilter((CellAfterWriteFilter) filter);
                continue;
            }
            throw new IllegalArgumentException("写入Table不支持的过滤器类型: " + filter.getClass());
        }
        return this;
    }

    public StandardWriter addFilter(String sheet, String pointKey, Collection<Filter> filters) {
        Check.notEmpty(filters);
        if (!sheetConfigMap.containsKey(sheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + sheet);
        }
        MySheet mySheetConfig = sheetConfigMap.get(sheet);
        Point point = mySheetConfig.findPoint(pointKey);

        for (Filter filter : filters) {
            if (filter instanceof CellBeforeWriteFilter) {
                point.addBeforeWriteFilter(filter);
                continue;
            }
            if (filter instanceof CellAfterWriteFilter) {
                point.addAfterWriteFilter(filter);
                continue;
            }
            throw new IllegalArgumentException("写入Point不支持的过滤器类型: " + filter.getClass());
        }
        return this;
    }

    public StandardWriter execute() {
        Excel excel = new Excel();
        excel.setOutput(output);
        excel.setMySheets(new ArrayList<>(sheetConfigMap.values()));
        // 存在模板，优先使用模板
        Workbook workbook;
        if (null != template) {
            try {
                workbook = WorkbookFactory.create(template);
            } catch (IOException | InvalidFormatException e) {
                e.printStackTrace();
                throw new IllegalArgumentException("使用模板创建Workbook失败");
            }
        } else {
            workbook = useXlsx ? new XSSFWorkbook() : new HSSFWorkbook();
        }
        ExcelSupport support = new ExcelSupport();
        support.writeExcel(workbook, excel);
        return this;
    }

}
