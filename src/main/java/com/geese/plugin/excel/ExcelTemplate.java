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

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelTemplate implements ExcelOperations {
    public static final String TABLE_DATA_KEY = "tableData";
    public static final String POINT_DATA_KEY = "pointData";

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
            Assert.notNull(sheet, "根据名:[%s]未获取到Sheet", name);
            Object sheetData = ExcelOperationsProxyFactory.getProxy().readSheet(sheet, sheetMapping);
            returnValue.put(sheetMapping.getDataKey(), sheetData);
            // TODO: 获取workbook中所有的sheet    2017/3/11 workbook.iterator()
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
//                Object rowData = readRow(row, sheetMapping);
                // 过滤未通过
                if (EXCEL_NOT_FILTERED.equals(rowData)) {
                    returnValue.put(TABLE_DATA_KEY, tableData);
                    return returnValue;
                }
                tableData.add(rowData);
            }
            returnValue.put(TABLE_DATA_KEY, tableData);
        }

        // 读取Map数据（散列点）
        List<CellMapping> mapPoints = sheetMapping.getPoints();
        if (null != mapPoints && !mapPoints.isEmpty()) {
            Map mapPointsData = new LinkedHashMap();
            for (CellMapping mapPoint : mapPoints) {
                // 行号
                Integer rowNumber = mapPoint.getRowNumber();
                // 列号
                Integer columnNumber = mapPoint.getColumnNumber();
                Cell cell = sheet.getRow(rowNumber).getCell(columnNumber);
                Object cellData = readCell(cell, mapPoint);
                mapPointsData.put(mapPoint.getDataKey(), cellData);
            }
            returnValue.put(POINT_DATA_KEY, mapPointsData);
        }

        return returnValue;
    }

    @Override
    public Object readRow(Row row, SheetMapping sheetMapping) {
        Assert.notNull(row, sheetMapping);
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
    public void write(Workbook workbook, ExcelMapping excelMapping) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void write(Sheet sheet, SheetMapping sheetMapping) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void write(Row row, SheetMapping sheetMapping, Object data) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void write(Cell cell, SheetMapping sheetMapping, Object data) {
        throw new UnsupportedOperationException();
    }
}
