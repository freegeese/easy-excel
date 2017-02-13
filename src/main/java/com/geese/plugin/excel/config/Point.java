package com.geese.plugin.excel.config;

/**
 * Cell 单元格配置信息
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 15:58
 * @sine 0.0.1
 */
public class Point extends Filterable {
    /**
     * 行号
     */
    private Integer x;

    /**
     * 列号
     */
    private Integer y;

    /**
     * 数据映射名称
     */
    private String key;

    /**
     * 数据
     */
    private Object data;

    /**
     * 关联的sheet配置信息，把point当做是sheet中的一个散列点
     */
    private MySheet mySheet;

    /**
     * 关联的table配置信息，把point当做是table中一行中的一列
     */
    private Table table;

    public Integer getX() {
        return x;
    }

    public void setX(Integer x) {
        this.x = x;
    }

    public Integer getY() {
        return y;
    }

    public void setY(Integer y) {
        this.y = y;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public MySheet getMySheet() {
        return mySheet;
    }

    public void setMySheet(MySheet mySheet) {
        this.mySheet = mySheet;
    }

    public Table getTable() {
        return table;
    }

    public void setTable(Table table) {
        this.table = table;
    }
}
