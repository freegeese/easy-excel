package com.geese.plugin.excel;

import com.geese.plugin.excel.core.ExcelHelper;
import com.geese.plugin.excel.mapping.CellMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import com.geese.plugin.excel.mapping.SheetMapping;
import com.geese.plugin.excel.util.Assert;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelTemplate implements ExcelOperations {
    static final String TABLE_DATA_KEY = "tableData";
    static final String POINT_DATA_KEY = "pointData";


    @Override
    public Object readExcel(Workbook workbook, ExcelMapping excelMapping) {
        Assert.notNull(workbook, excelMapping);
        List<SheetMapping> sheetMappings = excelMapping.getSheetMappings();
        Assert.notEmpty(sheetMappings);
        Map returnValue = new HashMap<>();
        for (SheetMapping sheetMapping : sheetMappings) {
            // 根据名称获取真实的Sheet
            String name = sheetMapping.getName();
            if (null != name) {
                Sheet sheet = workbook.getSheet(name);
                Assert.notNull(sheet, "根据名称:[%s]未获取到Sheet", name);
                Object sheetData = readSheet(sheet, sheetMapping);
                returnValue.put(sheetMapping.getTableDataKey(), sheetData);
                continue;
            }
            // 根据索引获取真实的Sheet
            Integer index = sheetMapping.getIndex();
            if (null != index) {
                Sheet sheet = workbook.getSheetAt(index);
                Assert.notNull(sheet, "根据索引:[%s]未获取到Sheet", index);
                Object sheetData = readSheet(sheet, sheetMapping);
                returnValue.put(sheetMapping.getTableDataKey(), sheetData);
                continue;
            }
            // TODO: 获取workbook中所有的sheet    2017/3/11 workbook.iterator()
            // 根据名称和索引均未找到Sheet
            throw new IllegalArgumentException("不能获取到Sheet, 根据SheetMapping的映射信息");
        }
        return returnValue;
    }

    @Override
    public Object readSheet(Sheet sheet, SheetMapping sheetMapping) {
        Assert.notNull(sheet, sheetMapping);

        Map returnValue = new HashMap<>();
        // 读取Table数据
        List<CellMapping> tableHeads = sheetMapping.getTableHeads();
        if (null != tableHeads && !tableHeads.isEmpty()) {
            // 开始行
            Integer startRow = sheetMapping.getStartRow();
            startRow = (null == startRow) ? 0 : startRow;
            // 结束行
            Integer endRow = sheetMapping.getEndRow();
            int lastRowNum = sheet.getLastRowNum();
            endRow = (null == endRow) ? lastRowNum : endRow;
            endRow = Math.min(endRow, lastRowNum);
            // 迭代行
            List tableData = new ArrayList();
            for (int i = startRow; i < endRow; i++) {
                Row row = sheet.getRow(i);
                Object rowData = readRow(row, sheetMapping);
                tableData.add(rowData);
            }
            returnValue.put(TABLE_DATA_KEY, tableData);
        }

        // 读取Map数据（散列点）
        List<CellMapping> mapPoints = sheetMapping.getPoints();
        if (null != mapPoints && !mapPoints.isEmpty()) {
            Map mapPointsData = new HashMap();
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

        Map returnValue = new HashMap<>();
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

    }

    @Override
    public void write(Sheet sheet, SheetMapping sheetMapping) {

    }

    @Override
    public void write(Row row, SheetMapping sheetMapping, Object data) {

    }

    @Override
    public void write(Cell cell, SheetMapping sheetMapping, Object data) {

    }
}
