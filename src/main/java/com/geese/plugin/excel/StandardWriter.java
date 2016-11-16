package com.geese.plugin.excel;

import com.geese.plugin.excel.config.ExcelConfig;
import com.geese.plugin.excel.core.ExcelHelper;
import com.geese.plugin.excel.core.ExcelSupport;
import com.geese.plugin.excel.core.OperationKey;
import com.geese.plugin.excel.util.Check;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.SheetConfig;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.filter.CellWriteFilter;
import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.filter.RowWriteFilter;
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
     * 写入Excel所需的Sheet配置信息
     */
    private Map<String, SheetConfig> sheetConfigMap;

    /**
     * 写入Excel时使用的模板
     */
    private InputStream template;

    public StandardWriter() {
        this.sheetConfigMap = new LinkedHashMap();
    }

    public static StandardWriter build(OutputStream output) {
        Check.notNull(output);
        StandardWriter writer = new StandardWriter();
        writer.output = output;
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
                SheetConfig sheetConfig = sheetConfigMap.get(sheet);
                if (null == sheetConfig) {
                    sheetConfig = new SheetConfig();
                    ExcelHelper.setSheet(sheet, sheetConfig);
                    sheetConfigMap.put(sheet, sheetConfig);
                }
                // [column row name]
                String[] rowColumnNames = keyDataMap.get(OperationKey.COLUMN).split(",");
                for (String rowColumnName : rowColumnNames) {
                    String[] items = rowColumnName.trim().split("-|\\s+");
                    Point point = new Point();
                    point.setX(Integer.valueOf(items[0]));
                    point.setY(Integer.valueOf(items[1]));
                    point.setKey(items[2]);
                    sheetConfig.addPoint(point);
                    point.setSheetConfig(sheetConfig);
                }
                continue;
            }

            // 列表插入 insert table
            Map<OperationKey, String> keyDataMap = ExcelHelper.insertKeyParse(insert);
            // into sheet
            String sheet = keyDataMap.get(OperationKey.INTO);
            SheetConfig sheetConfig = sheetConfigMap.get(sheet);
            if (null == sheetConfig) {
                sheetConfig = new SheetConfig();
                ExcelHelper.setSheet(sheet, sheetConfig);
                sheetConfigMap.put(sheet, sheetConfig);
            }

            // insert columns
            Table table = new Table();
            String[] columnNames = keyDataMap.get(OperationKey.COLUMN).split(",");
            for (String columnName : columnNames) {
                String[] columnWithName = columnName.trim().split("\\s+");
                Point point = new Point();
                point.setY(Integer.valueOf(columnWithName[0]));
                point.setKey(columnWithName[1]);
                table.addQueryPoint(point);
                point.setTable(table);
            }

            // limit
            if (keyDataMap.containsKey(OperationKey.LIMIT)) {
                String[] startWithSize = keyDataMap.get(OperationKey.LIMIT).split(",");
                table.setStartRow(Integer.valueOf(startWithSize[0].trim()));
                if (startWithSize.length > 1) {
                    table.setRowSize(Integer.valueOf(startWithSize[1].trim()));
                }
            }

            sheetConfig.addTable(table);
            table.setSheetConfig(sheetConfig);
        }
        return this;
    }

    /**
     * 添加Table数据
     *
     * @param toSheet
     * @param tableIndex
     * @param tableData
     * @return
     */
    public StandardWriter addData(String toSheet, Integer tableIndex, Collection tableData) {
        Check.notEmpty(toSheet, tableIndex, tableData);
        if (!sheetConfigMap.containsKey(toSheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + toSheet);
        }
        SheetConfig sheetConfig = sheetConfigMap.get(toSheet);
        Table table = sheetConfig.getTableList().get(tableIndex);
        table.setData(tableData);
        return this;
    }

    /**
     * 添加散列点数据
     *
     * @param toSheet
     * @param pointData
     * @return
     */
    public StandardWriter addData(String toSheet, Map pointData) {
        Check.notEmpty(toSheet, pointData);
        if (!sheetConfigMap.containsKey(toSheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + toSheet);
        }
        SheetConfig sheetConfig = sheetConfigMap.get(toSheet);
        sheetConfig.setPointData(pointData);
        return this;
    }

    /**
     * 添加过滤器到Sheet上
     *
     * @param toSheet
     * @param first
     * @param second
     * @param more
     * @return this
     */
    public StandardWriter addFilter(String toSheet, Filter first, Filter second, Filter... more) {
        Check.notNull(first, second);
        List<Filter> filters = new ArrayList<>();
        filters.add(first);
        filters.add(second);
        if (null != more) {
            filters.addAll(Arrays.asList(more));
        }
        return addFilter(toSheet, filters);
    }

    /**
     * 添加过滤器
     *
     * @param toSheet
     * @param filter
     * @return
     */
    public StandardWriter addFilter(String toSheet, Filter filter) {
        return addFilter(toSheet, Arrays.asList(filter));
    }

    /**
     * 添加过滤器
     *
     * @param toSheet
     * @param filters
     * @return
     */
    public StandardWriter addFilter(String toSheet, Filter[] filters) {
        return addFilter(toSheet, Arrays.asList(filters));
    }

    /**
     * 添加过滤器
     *
     * @param toSheet
     * @param filters
     * @return
     */
    public StandardWriter addFilter(String toSheet, Collection<Filter> filters) {
        Check.notEmpty(filters);
        if (!sheetConfigMap.containsKey(toSheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + toSheet);
        }
        SheetConfig sheetConfig = sheetConfigMap.get(toSheet);
        for (Filter filter : filters) {
            if (filter instanceof RowWriteFilter) {
                sheetConfig.addRowBeforeWriteFilter((RowWriteFilter) filter);
                continue;
            }
            if (filter instanceof CellWriteFilter) {
                sheetConfig.addCellBeforeWriteFilter((CellWriteFilter) filter);
                continue;
            }
            throw new IllegalArgumentException("写入 Excel 不支持的过滤器类型: " + filter.getClass());
        }
        return this;
    }

    /**
     * 执行 Excel 写操作
     *
     * @return
     */
    public StandardWriter execute() {
        return execute(true);
    }

    public StandardWriter execute(boolean useXlsx) {
        ExcelConfig excelConfig = new ExcelConfig();
        excelConfig.setOutput(output);
        excelConfig.setSheetConfigs(sheetConfigMap.values());
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
        support.writeExcel(workbook, excelConfig);
        return this;
    }

}
