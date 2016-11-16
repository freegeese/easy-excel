package com.geese.plugin.excel.config;

import com.geese.plugin.excel.filter.*;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Sheet 配置信息
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 15:59
 * @sine 0.0.1
 */
public class SheetConfig {
    /**
     * 映射真实sheet的索引
     */
    private Integer sheetIndex;

    /**
     * 映射真实sheet的名称
     */
    private String sheetName;

    /**
     * 读取行之前进行过滤
     */
    private FilterChain rowBeforeReadFilterChain;

    /**
     * 读取行之后进行过滤
     */
    private FilterChain rowAfterReadFilterChain;

    /**
     * 读取列之前进行过滤
     */
    private FilterChain cellBeforeReadFilterChain;

    /**
     * 读取列之后进行过滤
     */
    private FilterChain cellAfterReadFilterChain;

    /**
     * 写入行之前进行过滤
     */
    private FilterChain rowWriteFilterChain;

    /**
     * 写入列之前进行过滤
     */
    private FilterChain cellWriteFilterChain;

    /**
     * 映射sheet中的列表数据（线性数据）
     */
    private List<Table> tableList;

    /**
     * 映射sheet中的散列点（键值对数据）
     */
    private List<Point> pointList;

    /**
     * 散列点的数据源
     */
    private Map pointData;

    /**
     * 关联的Excel配置信息
     */
    private ExcelConfig excelConfig;

    public SheetConfig() {
        this.sheetIndex = 0;
    }

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public FilterChain getRowBeforeReadFilterChain() {
        return rowBeforeReadFilterChain;
    }

    public void setRowBeforeReadFilterChain(FilterChain rowBeforeReadFilterChain) {
        this.rowBeforeReadFilterChain = rowBeforeReadFilterChain;
    }

    public FilterChain getRowAfterReadFilterChain() {
        return rowAfterReadFilterChain;
    }

    public void setRowAfterReadFilterChain(FilterChain rowAfterReadFilterChain) {
        this.rowAfterReadFilterChain = rowAfterReadFilterChain;
    }

    public FilterChain getCellBeforeReadFilterChain() {
        return cellBeforeReadFilterChain;
    }

    public void setCellBeforeReadFilterChain(FilterChain cellBeforeReadFilterChain) {
        this.cellBeforeReadFilterChain = cellBeforeReadFilterChain;
    }

    public FilterChain getCellAfterReadFilterChain() {
        return cellAfterReadFilterChain;
    }

    public void setCellAfterReadFilterChain(FilterChain cellAfterReadFilterChain) {
        this.cellAfterReadFilterChain = cellAfterReadFilterChain;
    }

    public List<Table> getTableList() {
        return tableList;
    }

    public void setTableList(List<Table> tableList) {
        this.tableList = tableList;
    }

    public List<Point> getPointList() {
        return pointList;
    }

    public void setPointList(List<Point> pointList) {
        this.pointList = pointList;
    }

    public ExcelConfig getExcelConfig() {
        return excelConfig;
    }

    public void setExcelConfig(ExcelConfig excelConfig) {
        this.excelConfig = excelConfig;
    }

    public FilterChain getRowWriteFilterChain() {
        return rowWriteFilterChain;
    }

    public void setRowWriteFilterChain(FilterChain rowWriteFilterChain) {
        this.rowWriteFilterChain = rowWriteFilterChain;
    }

    public FilterChain getCellWriteFilterChain() {
        return cellWriteFilterChain;
    }

    public void setCellWriteFilterChain(FilterChain cellWriteFilterChain) {
        this.cellWriteFilterChain = cellWriteFilterChain;
    }

    public Map getPointData() {
        return pointData;
    }

    public void setPointData(Map pointData) {
        this.pointData = pointData;
    }


    public SheetConfig addTable(Table table) {
        if (null == tableList) {
            tableList = new ArrayList<>();
        }
        tableList.add(table);
        return this;
    }

    public SheetConfig addPoint(Point point) {
        if (null == pointList) {
            pointList = new ArrayList<>();
        }
        pointList.add(point);
        return this;
    }

    public SheetConfig addRowBeforeReadFilter(RowBeforeReadFilter filter) {
        if (null == rowBeforeReadFilterChain) {
            rowBeforeReadFilterChain = new FilterChain();
        }
        rowBeforeReadFilterChain.addFilter(filter);
        return this;
    }

    public SheetConfig addRowAfterReadFilter(RowAfterReadFilter filter) {
        if (null == rowAfterReadFilterChain) {
            rowAfterReadFilterChain = new FilterChain();
        }
        rowAfterReadFilterChain.addFilter(filter);
        return this;
    }

    public SheetConfig addCellBeforeReadFilter(CellBeforeReadFilter filter) {
        if (null == cellBeforeReadFilterChain) {
            cellBeforeReadFilterChain = new FilterChain();
        }
        cellBeforeReadFilterChain.addFilter(filter);
        return this;
    }

    public SheetConfig addCellAfterReadFilter(CellAfterReadFilter filter) {
        if (null == cellAfterReadFilterChain) {
            cellAfterReadFilterChain = new FilterChain();
        }
        cellAfterReadFilterChain.addFilter(filter);
        return this;
    }

    public SheetConfig addRowBeforeWriteFilter(RowWriteFilter filter) {
        if (null == rowWriteFilterChain) {
            rowWriteFilterChain = new FilterChain();
        }
        rowWriteFilterChain.addFilter(filter);
        return this;
    }

    public SheetConfig addCellBeforeWriteFilter(CellWriteFilter filter) {
        if (null == cellWriteFilterChain) {
            cellWriteFilterChain = new FilterChain();
        }
        cellWriteFilterChain.addFilter(filter);
        return this;
    }

}
