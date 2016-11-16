package com.geese.plugin.excel;

import com.geese.plugin.excel.config.ExcelConfig;
import com.geese.plugin.excel.core.ExcelHelper;
import com.geese.plugin.excel.core.OperationKey;
import com.geese.plugin.excel.filter.*;
import com.geese.plugin.excel.util.Check;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.SheetConfig;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.core.ExcelSupport;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * 标准Excel读取操作
 * <p> 通过使用模拟SQL查询语句来进行excel的读取操作，提供了比较全面的excel读取操作<br>
 * <ul>
 * <li>
 * 一个sheet一个table查询
 * {@code select("0 name, 1 age, 2 idCard from Sheet1 limit 0,10")}
 * </li>
 * <li>
 * 一个sheet多个table查询
 * {@code select(
 * "0 name, 1 age, 2 idCard from Sheet1 limit 0,10",
 * "4 phone, 5 email from Sheet1 limit 20, 10"
 * )}
 * </li>
 * <li>
 * 多个sheet多个table查询
 * {@code select(
 * "0 name, 1 age, 2 idCard from Sheet1 limit 0,10",
 * "4 phone, 5 email from 1 limit 20, 10"
 * )}
 * </li>
 * <li>
 * 多个sheet多个table查询多个散列点查询
 * {@code select(
 * "0 name, 1 age, 2 idCard from Sheet1 limit 0,10",
 * "1 phone, 2 email from 1 limit 20, 10",
 * "{0-4 birth, 1-5 color from Sheet1}",
 * "{0-2 foo, 1-2 bar from 1}"
 * )}
 * </li>
 * <p>
 * </ul>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 17:10
 * @sine 0.0.1
 */
public class StandardReader {

    /**
     * 读取excel所需的输入源
     */
    private InputStream input;

    /**
     * 读取excel所需的sheet配置信息
     */
    private Map<String, SheetConfig> sheetConfigMap;

    public StandardReader() {
        this.sheetConfigMap = new LinkedHashMap();
    }

    public static StandardReader build(InputStream input) {
        Check.notNull(input);
        StandardReader reader = new StandardReader();
        reader.input = input;
        return reader;
    }

    /**
     * 接受客户端输入的查询语句，解析之后放入到SheetConfig里面
     *
     * @param firstQuery
     * @param more
     * @return
     */
    public StandardReader select(String firstQuery, String... more) {
        Check.notEmpty(firstQuery);
        List<String> queryList = new ArrayList();
        queryList.add(firstQuery);
        if (null != more) {
            queryList.addAll(Arrays.asList(more));
        }

        for (String query : queryList) {
            query = query.trim();
            // 散列点查询 Point
            if (query.matches("^\\{.+\\}$")) {
                query = query.replaceAll("\\{|\\}", "");
                Map<OperationKey, String> keyDataMap = ExcelHelper.selectKeyParse(query);
                String sheet = keyDataMap.get(OperationKey.FROM);
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
            // 列表查询 Table
            Map<OperationKey, String> keyDataMap = ExcelHelper.selectKeyParse(query);
            // from sheet
            String sheet = keyDataMap.get(OperationKey.FROM);
            SheetConfig sheetConfig = sheetConfigMap.get(sheet);
            if (null == sheetConfig) {
                sheetConfig = new SheetConfig();
                ExcelHelper.setSheet(sheet, sheetConfig);
                sheetConfigMap.put(sheet, sheetConfig);
            }

            // select columns
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

            // where
            if (keyDataMap.containsKey(OperationKey.WHERE)) {
                table.setWhere(keyDataMap.get(OperationKey.WHERE));
            }

            // limit
            if (keyDataMap.containsKey(OperationKey.LIMIT)) {
                String[] startWithSize = keyDataMap.get(OperationKey.LIMIT).replaceAll("\\s+","").split(",");
                table.setStartRow(Integer.valueOf(startWithSize[0]));
                if (startWithSize.length > 1) {
                    table.setRowSize(Integer.valueOf(startWithSize[1]));
                }
            }
            sheetConfig.addTable(table);
            table.setSheetConfig(sheetConfig);
        }
        return this;
    }


    /**
     * 添加命名的参数一个sheet
     *
     * @param toSheet
     * @param tableIndex
     * @param namedParameterMap
     * @return this
     */
    public StandardReader addParameter(String toSheet, Integer tableIndex, Map<String, Object> namedParameterMap) {
        Check.notEmpty(namedParameterMap, toSheet, tableIndex);
        if (!sheetConfigMap.containsKey(toSheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + toSheet);
        }
        SheetConfig sheetConfig = sheetConfigMap.get(toSheet);
        Table table = sheetConfig.getTableList().get(tableIndex);
        table.setWhereParameter(namedParameterMap);
        return this;
    }

    /**
     * 添加占位符参数
     *
     * @param placeholderValues
     * @return this
     */
    public StandardReader addParameter(String toSheet, Integer tableIndex, Object[] placeholderValues) {
        return addParameter(toSheet, tableIndex, Arrays.asList(placeholderValues));
    }

    /**
     * 添加占位符参数
     *
     * @param placeholderValues
     * @return this
     */
    public StandardReader addParameter(String toSheet, Integer tableIndex, Collection placeholderValues) {
        Check.notEmpty(toSheet, tableIndex, placeholderValues);
        if (!sheetConfigMap.containsKey(toSheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + toSheet);
        }
        SheetConfig sheetConfig = sheetConfigMap.get(toSheet);
        Table table = sheetConfig.getTableList().get(tableIndex);
        Map placeholderValuesMap = new LinkedHashMap();
        int index = 0;
        for (Object placeholderValue : placeholderValues) {
            placeholderValuesMap.put(index++, placeholderValue);
        }
        table.setWhereParameter(placeholderValuesMap);
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
    public StandardReader addFilter(String toSheet, Filter first, Filter second, Filter... more) {
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
    public StandardReader addFilter(String toSheet, Filter filter) {
        return addFilter(toSheet, Arrays.asList(filter));
    }

    /**
     * 添加过滤器
     *
     * @param toSheet
     * @param filters
     * @return
     */
    public StandardReader addFilter(String toSheet, Filter[] filters) {
        return addFilter(toSheet, Arrays.asList(filters));
    }

    /**
     * 添加过滤器
     *
     * @param toSheet
     * @param filters
     * @return
     */
    public StandardReader addFilter(String toSheet, Collection<Filter> filters) {
        Check.notEmpty(filters);
        if (!sheetConfigMap.containsKey(toSheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + toSheet);
        }
        SheetConfig sheetConfig = sheetConfigMap.get(toSheet);
        for (Filter filter : filters) {
            if (filter instanceof RowBeforeReadFilter) {
                sheetConfig.addRowBeforeReadFilter((RowBeforeReadFilter) filter);
                continue;
            }
            if (filter instanceof RowAfterReadFilter) {
                sheetConfig.addRowAfterReadFilter((RowAfterReadFilter) filter);
                continue;
            }
            if (filter instanceof CellBeforeReadFilter) {
                sheetConfig.addCellBeforeReadFilter((CellBeforeReadFilter) filter);
                continue;
            }
            if (filter instanceof CellAfterReadFilter) {
                sheetConfig.addCellAfterReadFilter((CellAfterReadFilter) filter);
                continue;
            }
            throw new IllegalArgumentException("读取Excel不支持的过滤器类型: " + filter.getClass());
        }
        return this;
    }

    /**
     * 执行读Excel的操作
     *
     * @return this
     */
    public Object execute() {
        ExcelConfig excelConfig = new ExcelConfig();
        excelConfig.setInput(input);
        excelConfig.setSheetConfigs(sheetConfigMap.values());
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(input);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        ExcelSupport support = new ExcelSupport();
        return support.readExcel(workbook, excelConfig);
    }
}
