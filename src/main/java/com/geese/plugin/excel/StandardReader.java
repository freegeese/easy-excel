package com.geese.plugin.excel;

import com.geese.plugin.excel.config.Excel;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.Sheet;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.core.ExcelHelper;
import com.geese.plugin.excel.core.ExcelSupport;
import com.geese.plugin.excel.core.OperationKey;
import com.geese.plugin.excel.filter.*;
import com.geese.plugin.excel.util.Check;
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
    private Map<String, Sheet> sheetConfigMap;

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
                Sheet sheat = sheetConfigMap.get(sheet);
                if (null == sheat) {
                    sheat = new Sheet();
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
                    point.setSheet(sheat);
                }
                continue;
            }
            // 列表查询 Table
            Map<OperationKey, String> keyDataMap = ExcelHelper.selectKeyParse(query);
            // from sheet
            String sheet = keyDataMap.get(OperationKey.FROM);
            Sheet sheat = sheetConfigMap.get(sheet);
            if (null == sheat) {
                sheat = new Sheet();
                ExcelHelper.setSheet(sheet, sheat);
                sheetConfigMap.put(sheet, sheat);
            }

            // select columns
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
                String[] startWithSize = keyDataMap.get(OperationKey.LIMIT).replaceAll("\\s+", "").split(",");
                table.setStartRow(Integer.valueOf(startWithSize[0]));
                if (startWithSize.length > 1) {
                    table.setEndRow(table.getStartRow() + Integer.valueOf(startWithSize[1]));
                }
            }
            sheat.addTable(table);
            table.setSheet(sheat);
        }
        return this;
    }

    public StandardReader addFilter(String sheet, Integer tableIndex, Filter filter, Filter... more) {
        List<Filter> filters = new ArrayList();
        filters.add(filter);
        if (null == more) {
            filters.addAll(Arrays.asList(more));
        }
        return addFilter(sheet, tableIndex, filters);
    }

    public StandardReader addFilter(String sheet, Integer tableIndex, Filter[] filters) {
        return addFilter(sheet, tableIndex, Arrays.asList(filters));
    }

    public StandardReader addFilter(String sheet, Integer tableIndex, Collection<Filter> filters) {
        Check.notEmpty(filters);
        if (!sheetConfigMap.containsKey(sheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + sheet);
        }
        Sheet sheetConfig = sheetConfigMap.get(sheet);
        Table table = sheetConfig.getTables().get(tableIndex);

        for (Filter filter : filters) {
            if (filter instanceof RowBeforeReadFilter) {
                table.addRowBeforeReadFilter((RowBeforeReadFilter) filter);
                continue;
            }
            if (filter instanceof RowAfterReadFilter) {
                table.addRowAfterReadFilter((RowAfterReadFilter) filter);
                continue;
            }
            if (filter instanceof CellBeforeReadFilter) {
                table.addCellBeforeReadFilter((CellBeforeReadFilter) filter);
                continue;
            }
            if (filter instanceof CellAfterReadFilter) {
                table.addCellAfterReadFilter((CellAfterReadFilter) filter);
                continue;
            }
            throw new IllegalArgumentException("读取Table不支持的过滤器类型: " + filter.getClass());
        }
        return this;
    }

    public StandardReader addFilter(String sheet, String pointKey, Collection<Filter> filters) {
        Check.notEmpty(filters);
        if (!sheetConfigMap.containsKey(sheet)) {
            throw new IllegalArgumentException("不存在的sheet : " + sheet);
        }
        Sheet sheetConfig = sheetConfigMap.get(sheet);
        Point point = sheetConfig.findPoint(pointKey);

        for (Filter filter : filters) {
            if (filter instanceof CellBeforeReadFilter) {
                point.addBeforeReadFilter(filter);
                continue;
            }
            if (filter instanceof CellAfterReadFilter) {
                point.addAfterReadFilter(filter);
                continue;
            }
            throw new IllegalArgumentException("读取Point不支持的过滤器类型: " + filter.getClass());
        }
        return this;
    }

    /**
     * 执行读Excel的操作
     *
     * @return this
     */
    public Object execute() {
        Excel excel = new Excel();
        excel.setInput(input);
        excel.setSheets(new ArrayList<Sheet>(sheetConfigMap.values()));
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(input);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        ExcelSupport support = new ExcelSupport();
        return support.readExcel(workbook, excel);
    }
}
