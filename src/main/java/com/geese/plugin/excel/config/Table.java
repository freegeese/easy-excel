package com.geese.plugin.excel.config;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * sheet中数据表格的配置信息
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:04
 * @sine 0.0.1
 */
public class Table {
    /**
     * 映射sheet中一个列表的头部单元格
     */
    private List<Point> headPointList;

    /**
     * 对table的过滤条件
     */
    private String where;
    /**
     * 对table过滤条件的所需的参数
     */
    private Map whereParameter;

    /**
     * 读取sheet的开始行
     */
    private Integer startRow;

    /**
     * 读取sheet的行数
     */
    private Integer rowSize;

    /**
     * 读取sheet的开始行列
     */
    private Integer startColumn;

    /**
     * 读取sheet的列间隔
     */
    private Integer columnStep;

    /**
     * table的数据源
     */
    private Collection data;

    /**
     * table配置关联的sheet配置信息
     */
    private SheetConfig sheetConfig;

    public Table() {
        this.startRow = 0;
        this.startColumn = 0;
        this.columnStep = 0;
    }

    public List<Point> getHeadPointList() {
        return headPointList;
    }

    public void setHeadPointList(List<Point> headPointList) {
        this.headPointList = headPointList;
    }

    public String getWhere() {
        return where;
    }

    public void setWhere(String where) {
        this.where = where;
    }

    public Map getWhereParameter() {
        return whereParameter;
    }

    public void setWhereParameter(Map whereParameter) {
        this.whereParameter = whereParameter;
    }

    public Integer getStartRow() {
        return startRow;
    }

    public void setStartRow(Integer startRow) {
        this.startRow = startRow;
    }

    public Integer getRowSize() {
        return rowSize;
    }

    public void setRowSize(Integer rowSize) {
        this.rowSize = rowSize;
    }

    public Integer getStartColumn() {
        return startColumn;
    }

    public void setStartColumn(Integer startColumn) {
        this.startColumn = startColumn;
    }

    public Integer getColumnStep() {
        return columnStep;
    }

    public void setColumnStep(Integer columnStep) {
        this.columnStep = columnStep;
    }

    public Collection getData() {
        return data;
    }

    public void setData(Collection data) {
        this.data = data;
    }

    public SheetConfig getSheetConfig() {
        return sheetConfig;
    }

    public void setSheetConfig(SheetConfig sheetConfig) {
        this.sheetConfig = sheetConfig;
    }

    public Table addQueryPoint(Point point) {
        if (null == headPointList) {
            headPointList = new ArrayList<>();
        }
        headPointList.add(point);
        return this;
    }
}
