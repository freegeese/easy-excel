package com.geese.plugin.excel.core;

import com.geese.plugin.excel.config.ExcelConfig;
import com.geese.plugin.excel.util.EmptyUtils;
import com.geese.plugin.excel.config.Point;
import com.geese.plugin.excel.config.SheetConfig;
import com.geese.plugin.excel.config.Table;
import com.geese.plugin.excel.filter.FilterChain;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * Excel 操作具体实现
 */
public class ExcelSupport implements ExcelOperation {
    @Override
    public Object readExcel(Workbook workbook, ExcelConfig config) {
        Collection<SheetConfig> sheetConfigList = config.getSheetConfigs();
        Map excelData = new HashMap();
        for (SheetConfig sheetConfig : sheetConfigList) {
            Sheet sheet;
            // 根据名称读取sheet
            String sheetName = sheetConfig.getSheetName();
            if (null != sheetName) {
                sheet = workbook.getSheet(sheetName);
                if (null == sheet) {
                    throw new IllegalArgumentException("根据sheet名称：" + sheetName + "，没有找到对应的sheet");
                }
                excelData.put(sheetName, readSheet(sheet, sheetConfig));
                continue;
            }
            // 根据下标读取sheet
            Integer sheetIndex = sheetConfig.getSheetIndex();
            if (null != sheetIndex) {
                sheet = workbook.getSheet(String.valueOf(sheetIndex));
                if (null == sheet) {
                    sheet = workbook.getSheetAt(sheetIndex);
                }
                excelData.put(sheetIndex, readSheet(sheet, sheetConfig));
                continue;
            }
            // 使用默认激活
            sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
            if (null == sheet) {
                throw new IllegalArgumentException("找到可读取的sheet");
            }
            excelData.put(sheetIndex, readSheet(sheet, sheetConfig));
        }
        return excelData;
    }

    @Override
    public void writeExcel(Workbook workbook, ExcelConfig config) {
        Collection<SheetConfig> sheetConfigList = config.getSheetConfigs();
        for (SheetConfig sheetConfig : sheetConfigList) {
            Sheet sheet;
            // 根据配置的名称获取sheet
            String sheetName = sheetConfig.getSheetName();
            if (null != sheetName) {
                sheet = workbook.getSheet(sheetName);
                // 不存在则根据名称创建
                if (null == sheet) {
                    sheet = workbook.createSheet(sheetName);
                }
                writeSheet(sheet, sheetConfig);
                continue;
            }

            // 根据配置的下标获取sheet
            Integer sheetIndex = sheetConfig.getSheetIndex();
            if (null != sheetIndex) {
                // 未指定模板，根据指定的下标新创建一个sheet
                if (workbook.getNumberOfSheets() == 0) {
                    sheet = workbook.createSheet(String.valueOf(sheetIndex));
                    writeSheet(sheet, sheetConfig);
                    continue;
                }
                // 指定模板，优先根据名称获取，其次再根据下标获取
                sheet = workbook.getSheet(String.valueOf(sheetIndex));
                if (null != sheet) {
                    writeSheet(sheet, sheetConfig);
                    continue;
                }
                sheet = workbook.getSheetAt(sheetIndex);
                if (null != sheet) {
                    writeSheet(sheet, sheetConfig);
                    continue;
                }
            }

            // 未指定名称和下标，并且是新创建的Workbook
            if (workbook.getNumberOfSheets() == 0) {
                sheet = workbook.createSheet();
                writeSheet(sheet, sheetConfig);
                continue;
            }
            // 未指定名称和下标，并且是模板，就用默认激活的
            sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
            if (null != sheet) {
                writeSheet(sheet, sheetConfig);
                continue;
            }
            throw new IllegalArgumentException("未找到可写入的sheet");
        }
        try {
            OutputStream output = config.getOutput();
            workbook.write(output);
            output.flush();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public Object readSheet(Sheet sheet, SheetConfig config) {
        // Table 处理
        Map sheetData = new HashMap();
        List<Table> tableList = config.getTableList();
        if (EmptyUtils.notEmpty(tableList)) {
            List<Object> tableDataList = new ArrayList<>();
            for (Table table : tableList) {
                tableDataList.add(readTable(sheet, table, config));
            }
            sheetData.put("tableDataList", tableDataList);
        }
        // Point 处理
        List<Point> points = config.getPointList();
        if (EmptyUtils.notEmpty(points)) {
            Map pointDataMap = new LinkedHashMap();
            for (Point point : points) {
                Object pointData = readCell(sheet.getRow(point.getX()).getCell(point.getY()), point);
                pointDataMap.put(point.getKey(), pointData);
            }
            sheetData.put("pointDataMap", pointDataMap);
        }
        return sheetData;
    }

    @Override
    public void writeSheet(Sheet sheet, SheetConfig config) {
        // Table 处理
        List<Table> tableList = config.getTableList();
        if (EmptyUtils.notEmpty(tableList)) {
            for (Table table : tableList) {
                writeTable(sheet, table, config);
            }
        }
        // Point 处理
        List<Point> points = config.getPointList();
        if (EmptyUtils.notEmpty(points)) {
            // TODO: 2016/11/13 将index point data -> named point data
            Map pointDataMap = config.getPointData();
            for (Point point : points) {
                String key = point.getKey();
                if (!pointDataMap.containsKey(key)) {
                    // TODO: 2016/11/13 没有找到point对应的数据
                    throw new IllegalArgumentException("没有找到point[key=" + key + "]对应的数据[" + pointDataMap + "]");
                }
                Row row = ExcelHelper.createRow(sheet, point.getX());
                Cell cell = ExcelHelper.createCell(row, point.getY());
                writeCell(cell, pointDataMap.get(key), point);
            }
        }
    }

    /**
     * 写入Table数据
     *
     * @param sheet
     * @param table
     * @param sheetConfig
     */
    private void writeTable(Sheet sheet, Table table, SheetConfig sheetConfig) {
        List tableData = new ArrayList<>(table.getData());
        // 开始行
        Integer startRow = table.getStartRow();
        startRow = (null == startRow) ? 0 : startRow;
        // 行数
        Integer rowSize = table.getRowSize();
        // 结束行
        Integer endRow = (null == rowSize) ? (startRow + tableData.size()) : (startRow + rowSize);
        // 写入行之前的过滤链
        FilterChain chain = sheetConfig.getRowWriteFilterChain();

        int dataIndex = 0;
        for (int i = startRow; i < endRow; i++) {
            Row row = ExcelHelper.createRow(sheet, i);
            Object rowData = tableData.get(dataIndex++);
            // 写row之前过滤
            if (null != chain) {
                chain.doFilter(row, rowData, table);
            }
            writeRow(row, rowData, table);
        }
    }

    /**
     * 读取Table数据
     *
     * @param sheet
     * @param table
     * @param sheetConfig
     * @return
     */
    private List readTable(Sheet sheet, Table table, SheetConfig sheetConfig) {
        // 开始行
        Integer startRow = table.getStartRow();
        // 行数
        Integer rowSize = table.getRowSize();
        // 结束行
        Integer endRow = (null == rowSize) ? sheet.getLastRowNum() : (startRow + rowSize - 1);
        // 读取行之前的过滤链
        FilterChain rowBeforeReadFilterChain = sheetConfig.getRowBeforeReadFilterChain();
        // 读取行之后的过滤链
        FilterChain rowAfterReadFilterChain = sheetConfig.getRowAfterReadFilterChain();
        List tableData = new ArrayList();
        for (int i = startRow; i <= endRow; i++) {
            Row row = sheet.getRow(i);
            // TODO 异常处理 row == null
            if (null == row) {
                continue;
            }

            // 读row之前过滤
            if (null != rowBeforeReadFilterChain) {
                rowBeforeReadFilterChain.doFilter(row, null, table);
            }

            // 读row
            Object rowData = readRow(row, table);

            // where 过滤
            String where = table.getWhere();
            if (EmptyUtils.notEmpty(where)) {
                if (rowData instanceof Map) {
                    if (!ExcelHelper.whereFilter(where, (Map) rowData, table.getWhereParameter())) {
                        continue;
                    }
                }
            }

            // 读row之后过滤
            if (null != rowAfterReadFilterChain) {
                rowAfterReadFilterChain.doFilter(row, rowData, table);
            }
            tableData.add(rowData);
        }
        return tableData;
    }

    @Override
    public Object readRow(Row row, Table table) {
        SheetConfig sheetConfig = table.getSheetConfig();
        FilterChain cellBeforeReadFilterChain = sheetConfig.getCellBeforeReadFilterChain();
        FilterChain cellAfterReadFilterChain = sheetConfig.getCellAfterReadFilterChain();

        Map rowData = new HashMap();
        for (Point point : table.getHeadPointList()) {
            // TODO 异常处理 cell == null
            Cell cell = row.getCell(point.getY());
            if (null == cell) {
                continue;
            }
            // 读cell之前过滤
            if (null != cellBeforeReadFilterChain) {
                cellBeforeReadFilterChain.doFilter(cell, null, point);
            }
            Object cellValue = readCell(cell, point);
            // 读cell之后过滤
            if (null != cellAfterReadFilterChain) {
                cellAfterReadFilterChain.doFilter(cell, cellValue, point);
            }
            rowData.put(point.getKey(), cellValue);
        }
        return rowData;
    }

    @Override
    public void writeRow(Row row, Object value, Table table) {
        // 键值对类型数据
        if (Map.class.isAssignableFrom(value.getClass())) {
            FilterChain chain = table.getSheetConfig().getCellWriteFilterChain();
            Map rowValue = (Map) value;
            for (Point point : table.getHeadPointList()) {
                Cell cell = ExcelHelper.createCell(row, point.getY());
                Object cellValue = rowValue.get(point.getKey());
                // 写入列之前过滤
                if (null != chain) {
                    chain.doFilter(cell, cellValue, point);
                }
                writeCell(cell, cellValue, point);
            }
            return;
        }
        throw new IllegalArgumentException("不支持的值类型处理：" + value.getClass());
    }

    @Override
    public Object readCell(Cell cell, Point point) {
        return ExcelHelper.getCellValue(cell);
        // TODO 格式化从单元格获取到的值
    }

    @Override
    public void writeCell(Cell cell, Object value, Point point) {
        // TODO 格式化处理设置到单元格的值
        ExcelHelper.setCellValue(cell, value);
    }
}
