package com.geese.plugin.excel.mapping;

import com.geese.plugin.excel.filter.Filterable;

import java.util.List;

/**
 * Created by Administrator on 2017/3/11.
 */
public class SheetMapping extends Filterable {

    // 名称
    private String name;
    // 索引
    private Integer index;
    // 数据Key
    private String dataKey;

    // 线性数据结构-------------------------------
    // 表头
    private List<CellMapping> tableHeads;
    // 开始行
    private Integer startRow;
    // 结束行
    private Integer endRow;
    // 表格数据
    private List tableData;

    // 散列数据结构-------------------------------
    // 散列点
    private List<CellMapping> points;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public List<CellMapping> getTableHeads() {
        return tableHeads;
    }

    public void setTableHeads(List<CellMapping> tableHeads) {
        this.tableHeads = tableHeads;
    }

    public Integer getStartRow() {
        return startRow;
    }

    public void setStartRow(Integer startRow) {
        this.startRow = startRow;
    }

    public Integer getEndRow() {
        return endRow;
    }

    public void setEndRow(Integer endRow) {
        this.endRow = endRow;
    }

    public List getTableData() {
        return tableData;
    }

    public void setTableData(List tableData) {
        this.tableData = tableData;
    }

    public List<CellMapping> getPoints() {
        return points;
    }

    public void setPoints(List<CellMapping> points) {
        this.points = points;
    }

    public String getDataKey() {
        return dataKey;
    }

    public void setDataKey(String dataKey) {
        this.dataKey = dataKey;
    }
}
