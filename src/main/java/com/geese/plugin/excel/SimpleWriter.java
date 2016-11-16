package com.geese.plugin.excel;

import com.geese.plugin.excel.util.Check;
import com.geese.plugin.excel.util.EmptyUtils;
import com.geese.plugin.excel.filter.Filter;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

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

    /**
     * 写入Excel数据所使用的模板
     */
    private InputStream template;

    /**
     * 写入Excel使用的insert语句
     */
    private String insert;

    /**
     * 写入到哪个sheet
     */
    private String sheet;

    /**
     * 写入的范围：开始行，行数
     */
    private String limit;

    /**
     * 写入时候所用到的过滤器
     */
    private Collection<Filter> filters;

    /**
     * 写入的数据源
     */
    private Collection data;

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
        this.insert = insert;
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

    public SimpleWriter addData(Collection data) {
        Check.notEmpty(data);
        this.data = data;
        return this;
    }

    public SimpleWriter addFilter(Filter first, Filter second, Filter... more) {
        Check.notNull(first, second);
        List<Filter> filters = new ArrayList<>();
        filters.add(first);
        filters.add(second);
        if (null != more) {
            filters.addAll(Arrays.asList(more));
        }
        return addFilter(filters);

    }

    public SimpleWriter addFilter(Filter filter) {
        return addFilter(Arrays.asList(filter));
    }

    public SimpleWriter addFilter(Filter[] filters) {
        return addFilter(Arrays.asList(filters));
    }

    public SimpleWriter addFilter(Collection<Filter> filters) {
        Check.notEmpty(filters);
        this.filters = filters;
        return this;
    }

    /**
     * 执行 Excel 写操作
     *
     * @return
     */
    public void execute() {
        execute(true);
    }

    public void execute(boolean useXlsx) {
        // 拼接query语句 insert + from + where + limit
        String insert = String.valueOf(this.insert);
        insert += " into " + this.sheet;
        if (null != limit) {
            insert += " limit " + this.limit;
        }
        // 构建一个标准的Reader
        StandardWriter writer = (null != template) ? StandardWriter.build(output, template) : StandardWriter.build(output);
        writer.insert(insert);
        // 数据
        if (EmptyUtils.notEmpty(data)) {
            writer.addData(this.sheet, 0, data);
        }
        // 过滤器
        if (EmptyUtils.notEmpty(filters)) {
            writer.addFilter(this.sheet, filters);
        }
        writer.execute(useXlsx);
    }
}
