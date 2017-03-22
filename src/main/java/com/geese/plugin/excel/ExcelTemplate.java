package com.geese.plugin.excel;

import com.geese.plugin.excel.mapping.CellMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import com.geese.plugin.excel.mapping.SheetMapping;
import com.geese.plugin.excel.util.Assert;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.*;
import java.util.logging.Logger;

/**
 * Excel操作接口模板
 */
public class ExcelTemplate implements ExcelOperations {
    private final static Logger logger = Logger.getLogger(ExcelTemplate.class.getName());

    // 本地的线程变量
    private static final ThreadLocal<Map> localContext = new ThreadLocal<Map>() {
        @Override
        protected Map initialValue() {
            return new LinkedHashMap<>();
        }
    };

    public static Map getContext() {
        return localContext.get();
    }

    public static void setContext(Map context) {
        localContext.set(context);
    }

    @Override
    public Object readExcel(Workbook workbook, ExcelMapping excelMapping) {
        Assert.notNull(workbook, excelMapping);
        Collection<SheetMapping> sheetMappings = excelMapping.getSheetMappings();
        Assert.notEmpty(sheetMappings);
        Map returnValue = new LinkedHashMap();
        Sheet sheet = null;
        for (SheetMapping sheetMapping : sheetMappings) {
            // 根据名称获取真实的Sheet
            String name = sheetMapping.getName();
            if (null != name) {
                sheet = workbook.getSheet(name);
                if (null == sheet) {
                    // 根据名称未获取到，再根据索引获取
                    Integer index = sheetMapping.getIndex();
                    if (null != index) {
                        sheet = workbook.getSheetAt(index);
                    }
                }
            }
            Assert.notNull(sheet, "根据名称:[%s]未获取到Sheet", name);
            Object sheetData = ExcelOperationsProxyFactory.getProxy().readSheet(sheet, sheetMapping);
            returnValue.put(sheetMapping.getDataKey(), sheetData);
        }
        return returnValue;
    }

    @Override
    public Object readSheet(Sheet sheet, SheetMapping sheetMapping) {
        Assert.notNull(sheet, sheetMapping);

        Map returnValue = new LinkedHashMap();
        // 读取Table数据
        List<CellMapping> tableHeads = sheetMapping.getTableHeads();
        if (null != tableHeads && !tableHeads.isEmpty()) {
            // 开始行
            Integer startRow = sheetMapping.getStartRow();
            startRow = (null == startRow) ? 0 : startRow;
            // 结束行
            Integer endRow = sheetMapping.getEndRow();
            int lastRowNum = sheet.getLastRowNum() + 1;
            endRow = (null == endRow) ? lastRowNum : Math.min(endRow, lastRowNum);
            // 迭代行
            List tableData = new ArrayList();
            for (int i = startRow; i < endRow; i++) {
                Row row = sheet.getRow(i);
                Object rowData = ExcelOperationsProxyFactory.getProxy().readRow(row, sheetMapping);
                // 过滤未通过
                if (EXCEL_NOT_PASS_FILTERED.equals(rowData)) {
                    returnValue.put(ExcelResult.TABLE_DATA_KEY, tableData);
                    return returnValue;
                }
                tableData.add(rowData);
            }
            returnValue.put(ExcelResult.TABLE_DATA_KEY, tableData);
        }

        // 读取Map数据（散列点）
        List<CellMapping> points = sheetMapping.getPoints();
        if (null != points && !points.isEmpty()) {
            Map mapPointsData = new LinkedHashMap();
            for (CellMapping point : points) {
                // 行号
                Integer rowNumber = point.getRowNumber();
                // 列号
                Integer columnNumber = point.getColumnNumber();
                Row row = sheet.getRow(rowNumber);
                Object cellData = null;
                if (null != row) {
                    Cell cell = row.getCell(columnNumber);
                    if (null != cell) {
                        cellData = readCell(cell, point);
                    }
                }
                mapPointsData.put(point.getDataKey(), cellData);

            }
            returnValue.put(ExcelResult.POINT_DATA_KEY, mapPointsData);
        }

        return returnValue;
    }

    @Override
    public Object readRow(Row row, SheetMapping sheetMapping) {
        if (null == row) {
            return null;
        }
        List<CellMapping> tableHeads = sheetMapping.getTableHeads();
        Assert.notEmpty(tableHeads);

        Map returnValue = new LinkedHashMap();
        for (CellMapping tableHead : tableHeads) {
            Integer columnNumber = tableHead.getColumnNumber();
            Cell cell = row.getCell(columnNumber);
            Object cellData = readCell(cell, tableHead);
            returnValue.put(tableHead.getDataKey(), cellData);
        }
        return returnValue;
    }

    @Override
    public Object readCell(Cell cell, CellMapping cellMapping) {
        return ExcelHelper.getCellValue(cell);
    }

    @Override
    public void writeExcel(Workbook workbook, ExcelMapping excelMapping) {
        Assert.notNull(workbook, excelMapping);
        Collection<SheetMapping> sheetMappings = excelMapping.getSheetMappings();
        Assert.notEmpty(sheetMappings);
        for (SheetMapping sheetMapping : sheetMappings) {
            Sheet sheet = getRawSheet(workbook, sheetMapping);
            if (null == sheet) {
                sheet = workbook.createSheet(sheetMapping.getName());
                sheetMapping.setIndex(workbook.getSheetIndex(sheet));
            }
            ExcelOperationsProxyFactory.getProxy().writeSheet(sheet, sheetMapping);
        }
    }

    /**
     * 获取真实的sheet
     *
     * @param workbook
     * @param sheetMapping
     * @return
     */
    private Sheet getRawSheet(Workbook workbook, SheetMapping sheetMapping) {
        Assert.notNull(workbook, sheetMapping);
        // 该Workbook没有sheet表格
        if (!workbook.iterator().hasNext()) {
            return null;
        }
        // 根据名称获取sheet
        String name = sheetMapping.getName();
        if (null != name) {
            Sheet sheet = workbook.getSheet(name);
            if (null != sheet) {
                sheetMapping.setIndex(workbook.getSheetIndex(sheet));
                return sheet;
            }
        }
        // 根据下标获取sheet
        Integer index = sheetMapping.getIndex();
        if (null != index) {
            Sheet sheet = workbook.getSheetAt(index);
            if (null != sheet) {
                sheetMapping.setName(sheet.getSheetName());
                return sheet;
            }
        }
        return null;
    }

    @Override
    public void writeSheet(Sheet sheet, SheetMapping sheetMapping) {
        // 处理表格数据
        List<Map> tableData = sheetMapping.getTableData();
        if (null != tableData && !tableData.isEmpty()) {
            // 开始行
            Integer startRow = sheetMapping.getStartRow();
            startRow = (null == startRow) ? 0 : startRow;
            // 行间隔
            Integer rowInterval = sheetMapping.getRowInterval();
            if (null == rowInterval || rowInterval <= 0) {
                rowInterval = 1;
            } else {
                rowInterval++;
            }
            for (Map rowData : tableData) {
                Row row = ExcelHelper.createRow(sheet, startRow);
                ExcelOperationsProxyFactory.getProxy().writeRow(row, sheetMapping, rowData);
                if (getContext().containsKey(EXCEL_NOT_PASS_FILTERED)) {
                    return;
                }
                startRow += rowInterval;
            }
        }
        // 处理散列数据
        List<CellMapping> points = sheetMapping.getPoints();
        if (null != points && !points.isEmpty()) {
            for (CellMapping point : points) {
                Integer rowNumber = point.getRowNumber();
                Integer columnNumber = point.getColumnNumber();
                Row row = ExcelHelper.createRow(sheet, rowNumber);
                Cell cell = ExcelHelper.createCell(row, columnNumber);
                ExcelOperationsProxyFactory.getProxy().writeCell(cell, sheetMapping, point.getData());
            }
        }
    }

    @Override
    public void writeRow(Row row, SheetMapping sheetMapping, Map data) {
        List<CellMapping> tableHeads = sheetMapping.getTableHeads();
        for (CellMapping head : tableHeads) {
            if (data.containsKey(head.getDataKey())) {
                Object cellData = data.get(head.getDataKey());
                Integer columnNumber = head.getColumnNumber();
                Cell cell = ExcelHelper.createCell(row, columnNumber);
                ExcelOperationsProxyFactory.getProxy().writeCell(cell, sheetMapping, cellData);
                continue;
            }
            logger.warning("数据：" + data + ", 不包含列：" + head.getDataKey());
        }
    }

    @Override
    public void writeCell(Cell cell, SheetMapping sheetMapping, Object data) {
        ExcelHelper.setCellValue(cell, data);
    }
}
