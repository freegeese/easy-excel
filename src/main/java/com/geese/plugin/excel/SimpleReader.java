package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.util.Check;
import com.geese.plugin.excel.util.EmptyUtils;

import java.io.InputStream;
import java.util.*;

/**
 * 简单Excel读取操作
 * <p>
 * 可以通过select(),from(),limit()几个简单的接口快速读取excel中的数据
 * </p>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:30
 * @sine 0.0.1
 */
public class SimpleReader {

    private InputStream input;

    private String tableQuery;

    private String pointQuery;

    private String sheet;

    private String limit;

    private List<Filter> tableFilters = new ArrayList<>();

    private Map<String, List<Filter>> pointFiltersMap = new HashMap<>();

    public SimpleReader() {
        this.sheet = "0";
    }

    public static SimpleReader build(InputStream input) {
        Check.notNull(input);
        SimpleReader reader = new SimpleReader();
        reader.input = input;
        return reader;
    }

    public SimpleReader select(String query) {
        Check.notEmpty(query);
        if (query.trim().startsWith("\\{")) {
            this.pointQuery = query;
        } else {
            this.tableQuery = query;
        }
        return this;
    }

    public SimpleReader select(String tableQuery, String pointQuery) {
        Check.notEmpty(tableQuery, pointQuery);
        this.tableQuery = tableQuery;
        this.pointQuery = pointQuery;
        return this;
    }

    public SimpleReader from(String sheet) {
        Check.notEmpty(sheet);
        this.sheet = sheet;
        return this;
    }

    public SimpleReader addFilter(Filter filter) {
        Check.notNull(filter);
        tableFilters.add(filter);
        return this;
    }


    public SimpleReader addFilter(Filter[] filters) {
        Check.notEmpty(filters);
        return addFilter(Arrays.asList(filters));
    }

    public SimpleReader addFilter(Collection<Filter> filters) {
        Check.notEmpty(filters);
        tableFilters.addAll(filters);
        return this;
    }

    public SimpleReader addFilter(String pointKey, Filter filter) {
        Check.notEmpty(pointKey, filter);
        return addFilter(pointKey, Arrays.asList(filter));
    }


    public SimpleReader addFilter(String pointKey, Filter[] filters) {
        Check.notEmpty(pointKey, filters);
        return addFilter(Arrays.asList(filters));
    }

    public SimpleReader addFilter(String pointKey, Collection<Filter> filters) {
        Check.notEmpty(pointKey, filters);
        List<Filter> values = pointFiltersMap.get(pointKey);
        if (null == values) {
            values = new ArrayList<>();
            pointFiltersMap.put(pointKey, values);
        }
        values.addAll(values);
        return this;
    }

    public SimpleReader limit(Integer startRow) {
        Check.notNull(startRow);
        this.limit = String.valueOf(startRow);
        return this;
    }

    public SimpleReader limit(Integer startRow, Integer size) {
        Check.notNull(startRow, size);
        this.limit = startRow + "," + size;
        return this;
    }

    public Collection execute() {
        // 拼接query语句 query + from + where + limit
        String tableQuery = null;
        if (EmptyUtils.notEmpty(this.tableQuery)) {
            tableQuery = String.valueOf(this.tableQuery);
            tableQuery += " from " + this.sheet;
            if (null != limit) {
                tableQuery += " limit " + this.limit;
            }
        }

        // 构建一个标准的Reader
        StandardReader reader = StandardReader.build(input);
        if (EmptyUtils.notEmpty(tableQuery) && EmptyUtils.notEmpty(this.pointQuery)) {
            reader.select(tableQuery, this.pointQuery);
        } else if (EmptyUtils.notEmpty(tableQuery)) {
            reader.select(tableQuery);
        } else {
            reader.select(this.pointQuery);
        }

        // 表格过滤器
        if (!tableFilters.isEmpty()) {
            reader.addFilter(this.sheet, 0, tableFilters);
        }
        // 散列点过滤器
        if (!pointFiltersMap.isEmpty()) {
            for (Map.Entry<String, List<Filter>> entry : pointFiltersMap.entrySet()) {
                reader.addFilter(this.sheet, entry.getKey(), entry.getValue());
            }
        }

        Map result = (Map) reader.execute();
        if (EmptyUtils.notEmpty(result)) {
            Map tableDataMap = (Map) result.values().iterator().next();
            if (EmptyUtils.notEmpty(tableDataMap)) {
                Collection tableDatas = (Collection) tableDataMap.values().iterator().next();
                if (EmptyUtils.notEmpty(tableDatas)) {
                    return (Collection) tableDatas.iterator().next();
                }
            }
        }
        return null;
    }

}
