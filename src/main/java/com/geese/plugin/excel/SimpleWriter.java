package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.util.Check;
import com.geese.plugin.excel.util.EmptyUtils;

import java.io.*;
import java.util.*;

/**
 * 简单Excel写入操作
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:57
 * @sine 0.0.1
 */
public class SimpleWriter {

    /**
     * 写入Excel数据之后，输出到哪里
     */
    private OutputStream output;

    private boolean useXlsx = true;

    /**
     * 写入Excel数据所使用的模板
     */
    private InputStream template;

    /**
     * 写入Excel使用的insert语句
     */
    private String tableInsert;

    private String pointInsert;

    /**
     * 写入到哪个sheet
     */
    private String sheet;

    /**
     * 写入的范围：开始行，行数
     */
    private String limit;

    /**
     * 写入的数据源
     */
    private List tableData;

    /**
     * 散列点数据
     */
    private Map pointData;

    /**
     * 表格过滤器
     */
    private List<Filter> tableFilters = new ArrayList<>();

    /**
     * 散列点过滤器
     */
    private Map<String, List<Filter>> pointFiltersMap = new HashMap<>();

    /**
     * 不存在模板的情况，创建一个新的sheet，名称为 Sheet1
     */
    public SimpleWriter() {
        this.sheet = "Sheet1";
    }

    public static SimpleWriter build(OutputStream output) {
        Check.notNull(output);
        SimpleWriter writer = new SimpleWriter();
        writer.output = output;
        return writer;
    }

    public static SimpleWriter build(OutputStream output, boolean useXlsx) {
        Check.notNull(output);
        SimpleWriter writer = new SimpleWriter();
        writer.output = output;
        writer.useXlsx = useXlsx;
        return writer;
    }

    public static SimpleWriter build(OutputStream output, File template) {
        Check.notNull(output, template);
        try {
            return build(output, new FileInputStream(template));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new IllegalArgumentException("模板文件不存在：" + template);
        }
    }

    public static SimpleWriter build(OutputStream output, InputStream template) {
        Check.notNull(output, template);
        SimpleWriter writer = new SimpleWriter();
        writer.output = output;
        writer.template = template;
        // 存在模板的时候，默认写入第0个sheet
        writer.sheet = "0";
        return writer;
    }

    public SimpleWriter insert(String insert) {
        Check.notEmpty(insert);
        this.tableInsert = insert;
        return this;
    }

    public SimpleWriter insert(String tableInsert, String pointInsert) {
        Check.notEmpty(tableInsert, pointInsert);
        this.tableInsert = tableInsert;
        this.pointInsert = pointInsert;
        return this;
    }

    public SimpleWriter into(String sheet) {
        Check.notEmpty(sheet);
        this.sheet = sheet;
        return this;
    }

    public SimpleWriter limit(Integer startRow) {
        Check.notNull(startRow);
        this.limit = String.valueOf(startRow);
        return this;
    }

    public SimpleWriter limit(Integer startRow, Integer size) {
        Check.notNull(startRow, size);
        this.limit = startRow + "," + size;
        return this;
    }

    public SimpleWriter addData(List tableData) {
        Check.notEmpty(tableData);
        this.tableData = tableData;
        return this;
    }

    public SimpleWriter addData(Map pointData) {
        Check.notEmpty(pointData);
        this.pointData = pointData;
        return this;
    }

    public SimpleWriter addData(List tableData, Map pointData) {
        Check.notEmpty(tableData, pointData);
        this.tableData = tableData;
        this.pointData = pointData;
        return this;
    }

    public SimpleWriter addFilter(Filter filter) {
        Check.notNull(filter);
        tableFilters.add(filter);
        return this;
    }


    public SimpleWriter addFilter(Filter[] filters) {
        Check.notEmpty(filters);
        return addFilter(Arrays.asList(filters));
    }

    public SimpleWriter addFilter(Collection<Filter> filters) {
        Check.notEmpty(filters);
        tableFilters.addAll(filters);
        return this;
    }

    public SimpleWriter addFilter(String pointKey, Filter filter) {
        Check.notEmpty(pointKey, filter);
        return addFilter(pointKey, Arrays.asList(filter));
    }


    public SimpleWriter addFilter(String pointKey, Filter[] filters) {
        Check.notEmpty(pointKey, filters);
        return addFilter(Arrays.asList(filters));
    }

    public SimpleWriter addFilter(String pointKey, Collection<Filter> filters) {
        Check.notEmpty(pointKey, filters);
        List<Filter> values = pointFiltersMap.get(pointKey);
        if (null == values) {
            values = new ArrayList<>();
            pointFiltersMap.put(pointKey, values);
        }
        values.addAll(values);
        return this;
    }

    public void execute() {
        // 拼接insert语句 insert into limit
        String insert = null;
        if (EmptyUtils.notEmpty(this.tableInsert)) {
            insert = String.valueOf(this.tableInsert);
            insert += " into " + this.sheet;
            if (null != limit) {
                insert += " limit " + this.limit;
            }
        }

        // 构建一个标准的Reader
        StandardWriter writer = (null != template) ? StandardWriter.build(output, template) : StandardWriter.build(output, useXlsx);

        if (EmptyUtils.notEmpty(insert) && EmptyUtils.notEmpty(pointInsert)) {
            writer.insert(insert, pointInsert);
        } else if (EmptyUtils.notEmpty(insert)) {
            writer.insert(insert);
        } else {
            writer.insert(pointInsert);
        }

        if (EmptyUtils.notEmpty(tableData)) {
            writer.addData(this.sheet, 0, tableData);
        }

        if (EmptyUtils.notEmpty(pointData)) {
            writer.addData(this.sheet, pointData);
        }

        if (EmptyUtils.notEmpty(tableFilters)) {
            writer.addFilter(this.sheet, 0, tableFilters);
        }

        if (EmptyUtils.notEmpty(pointFiltersMap)) {
            for (Map.Entry<String, List<Filter>> entry : pointFiltersMap.entrySet()) {
                writer.addFilter(this.sheet, entry.getKey(), entry.getValue());
            }
        }
        writer.execute();
    }
}
