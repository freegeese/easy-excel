package com.geese.plugin.excel;

import com.geese.plugin.excel.mapping.CellMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import com.geese.plugin.excel.mapping.SheetMapping;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * ExcelMapping 操作接口定义
 */
public interface ExcelOperations {
    /**
     * 读取Excel
     *
     * @param workbook
     * @param excelMapping
     * @return
     */
    Object readExcel(Workbook workbook, ExcelMapping excelMapping);

    /**
     * 读取Sheet
     *
     * @param sheet
     * @param sheetMapping
     * @return
     */
    Object readSheet(Sheet sheet, SheetMapping sheetMapping);

    /**
     * 读取Row
     *
     * @param row
     * @param sheetMapping
     * @return
     */
    Object readRow(Row row, SheetMapping sheetMapping);

    /**
     * 读取Cell
     *
     * @param cell
     * @param cellMapping
     * @return
     */
    Object readCell(Cell cell, CellMapping cellMapping);


    /**
     * 写入Excel
     *
     * @param workbook
     * @param excelMapping
     */
    void write(Workbook workbook, ExcelMapping excelMapping);

    /**
     * 写入Sheet
     *
     * @param sheet
     * @param sheetMapping
     */
    void write(Sheet sheet, SheetMapping sheetMapping);

    /**
     * 写入Row
     *
     * @param row
     * @param sheetMapping
     * @param data
     */
    void write(Row row, SheetMapping sheetMapping, Object data);

    /**
     * 写入Cell
     *
     * @param cell
     * @param sheetMapping
     * @param data
     */
    void write(Cell cell, SheetMapping sheetMapping, Object data);

}
