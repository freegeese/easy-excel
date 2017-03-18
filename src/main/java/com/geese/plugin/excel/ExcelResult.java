package com.geese.plugin.excel;

import java.util.Collection;
import java.util.Map;

/**
 * 读/写Excel后的结果封装
 */
public class ExcelResult {
    public static final String TABLE_DATA_KEY = "tableData";
    public static final String POINT_DATA_KEY = "pointData";

    // 读/写操作中的上下文
    private Map context;
    // 读写后的数据
    private Map data;

    public Map getContext() {
        return context;
    }

    public void setContext(Map context) {
        this.context = context;
    }

    public Map getData() {
        return data;
    }

    public void setData(Map data) {
        this.data = data;
    }

    /**
     * 获取指定sheet中的表格数据
     *
     * @param sheet
     * @return
     */
    public Collection getTableData(String sheet) {
        if (data.containsKey(sheet)) {
            return (Collection) ((Map) data.get(sheet)).get(TABLE_DATA_KEY);
        }
        return null;
    }

    /**
     * 获取指定sheet中的散点数据
     *
     * @param sheet
     * @return
     */
    public Map getPointData(String sheet) {
        if (data.containsKey(sheet)) {
            return (Map) ((Map) data.get(sheet)).get(POINT_DATA_KEY);
        }
        return null;
    }

    /**
     * 获取第一个sheet中的表格数据
     *
     * @return
     */
    public Collection getTableData() {
        if (!data.isEmpty()) {
            return getTableData(data.keySet().iterator().next().toString());
        }
        return null;
    }

    /**
     * 获取第一个sheet中的散列数据
     *
     * @return
     */
    public Map getPointData() {
        if (!data.isEmpty()) {
            return getPointData(data.keySet().iterator().next().toString());
        }
        return null;
    }
}
