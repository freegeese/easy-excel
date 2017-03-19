package com.geese.plugin.excel;

import com.geese.plugin.excel.mapping.CellMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import com.geese.plugin.excel.mapping.SheetMapping;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * ExcelMapping 操作接口定义
 */
public interface ExcelOperations {
    String EXCEL_NOT_PASS_FILTERED = "EXCEL_NOT_PASS_FILTERED";
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
    void writeExcel(Workbook workbook, ExcelMapping excelMapping);

    /**
     * 写入Sheet
     *
     * @param sheet
     * @param sheetMapping
     */
    void writeSheet(Sheet sheet, SheetMapping sheetMapping);

    /**
     * 写入Row
     *
     * @param row
     * @param sheetMapping
     * @param data
     */
    void writeRow(Row row, SheetMapping sheetMapping, Map data);

    /**
     * 写入Cell
     *
     * @param cell
     * @param sheetMapping
     * @param data
     */
    void writeCell(Cell cell, SheetMapping sheetMapping, Object data);

}
