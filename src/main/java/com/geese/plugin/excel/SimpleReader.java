package com.geese.plugin.excel;

import com.geese.plugin.excel.util.Check;
import com.geese.plugin.excel.util.EmptyUtils;
import com.geese.plugin.excel.filter.Filter;

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

    /**
     * 读取excel时候的输入源
     */
    private InputStream input;

    /**
     * 读取excel的查询语句
     */
    private String query;

    /**
     * 读取哪个sheet表格
     */
    private String sheet;

    /**
     * 读取excel使用的where条件
     */
    private String where;

    /**
     * 读取excel的限制读取范围：开始行，行数
     */
    private String limit;

    /**
     * 命名的参数数据
     */
    private Map namedParameterMap;

    /**
     * 读取excel使用的过滤器
     */
    private Collection<Filter> filters;

    /**
     * 如果没有指定读取的sheet，默认读取第0个
     */
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
        this.query = query;
        return this;
    }

    public SimpleReader from(String sheet) {
        Check.notEmpty(sheet);
        this.sheet = sheet;
        return this;
    }

    public SimpleReader where(String where) {
        Check.notEmpty(where);
        this.where = where;
        return this;
    }

    /**
     * 添加命名的参数一个sheet
     *
     * @param namedParameterMap
     * @return this
     */
    public SimpleReader addParameter(Map<String, Object> namedParameterMap) {
        Check.notEmpty(namedParameterMap);
        this.namedParameterMap = namedParameterMap;
        return this;
    }

    /**
     * 添加占位符参数
     *
     * @param placeholderValues
     * @return this
     */
    public SimpleReader addParameter(Object[] placeholderValues) {
        return addParameter(Arrays.asList(placeholderValues));
    }

    /**
     * 添加占位符参数
     *
     * @param placeholderValues
     * @return this
     */
    public SimpleReader addParameter(Collection placeholderValues) {
        Check.notEmpty(placeholderValues);
        Map placeholderValueMap = new LinkedHashMap();
        int index = 0;
        for (Object placeholderValue : placeholderValues) {
            placeholderValueMap.put(index++, placeholderValue);
        }
        this.namedParameterMap = placeholderValueMap;
        return this;
    }

    public SimpleReader addFilter(Filter first, Filter second, Filter... more) {
        Check.notNull(first, second);
        List<Filter> filters = new ArrayList<>();
        filters.add(first);
        filters.add(second);
        if (null != more) {
            filters.addAll(Arrays.asList(more));
        }
        return addFilter(filters);

    }

    public SimpleReader addFilter(Filter filter) {
        return addFilter(Arrays.asList(filter));
    }

    public SimpleReader addFilter(Filter[] filters) {
        return addFilter(Arrays.asList(filters));
    }

    public SimpleReader addFilter(Collection<Filter> filters) {
        Check.notEmpty(filters);
        this.filters = filters;
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
        String query = String.valueOf(this.query);
        query += " from " + this.sheet;
        if (null != where) {
            query += " where " + this.where;
        }
        if (null != limit) {
            query += " limit " + this.limit;
        }
        // 构建一个标准的Reader
        StandardReader reader = StandardReader.build(input).select(query);
        // 参数
        if (EmptyUtils.notEmpty(namedParameterMap)) {
            reader.addParameter(this.sheet, 0, namedParameterMap);
        }
        // 过滤器
        if (EmptyUtils.notEmpty(filters)) {
            reader.addFilter(this.sheet, filters);
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
