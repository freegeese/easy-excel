package com.geese.plugin.excel.config;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * MySheet 配置信息
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 15:59
 * @sine 0.0.1
 */
public class MySheet extends Filterable {
    /**
     * 映射真实sheet的索引
     */
    private Integer index;

    /**
     * 映射真实sheet的名称
     */
    private String name;

    /**
     * 映射sheet中的列表数据（线性数据）
     */
    private List<Table> tables;

    /**
     * 映射sheet中的散列点（键值对数据）
     */
    private List<Point> points;

    public MySheet addTable(Table table) {
        if (null == this.tables) {
            this.tables = new ArrayList<>();
        }
        this.tables.add(table);
        return this;
    }

    public MySheet addTables(Collection<Table> tables) {
        if (null == this.tables) {
            this.tables = new ArrayList<>();
        }
        this.tables.addAll(tables);
        return this;
    }

    public MySheet addPoint(Point point) {
        if (null == this.points) {
            this.points = new ArrayList<>();
        }
        this.points.add(point);
        return this;
    }

    public MySheet addPoints(Collection<Point> points) {
        if (null == this.points) {
            this.points = new ArrayList<>();
        }
        this.points.addAll(points);
        return this;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Table> getTables() {
        return tables;
    }

    public void setTables(List<Table> tables) {
        this.tables = tables;
    }

    public List<Point> getPoints() {
        return points;
    }

    public void setPoints(List<Point> points) {
        this.points = points;
    }

    public Point findPoint(String pointKey) {
        for (Point point : points) {
            if (pointKey.equals(point.getKey())) {
                return point;
            }
        }
        return null;
    }
}
