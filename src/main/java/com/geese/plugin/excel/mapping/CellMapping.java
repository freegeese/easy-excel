package com.geese.plugin.excel.mapping;

/**
 * Created by Administrator on 2017/3/11.
 */
public class CellMapping {

    // 单元格所在行
    private Integer columnNumber;
    // 单元格所在列
    private Integer rowNumber;
    // 数据名称映射
    private String dataKey;
    // 单元格数据
    private Object data;
    // 关联 Sheet
    private SheetMapping sheetMapping;

    public Integer getColumnNumber() {
        return columnNumber;
    }

    public void setColumnNumber(Integer columnNumber) {
        this.columnNumber = columnNumber;
    }

    public Integer getRowNumber() {
        return rowNumber;
    }

    public void setRowNumber(Integer rowNumber) {
        this.rowNumber = rowNumber;
    }

    public String getDataKey() {
        return dataKey;
    }

    public void setDataKey(String dataKey) {
        this.dataKey = dataKey;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public SheetMapping getSheetMapping() {
        return sheetMapping;
    }

    public void setSheetMapping(SheetMapping sheetMapping) {
        this.sheetMapping = sheetMapping;
    }
}
