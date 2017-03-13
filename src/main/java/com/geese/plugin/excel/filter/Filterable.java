package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.util.Assert;

import java.util.Collection;

/**
 * <p> 可过滤的 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:00
 * @sine 0.0.2
 */
public class Filterable {
    private FilterChain beforeReadFilterChain = new FilterChain();

    private FilterChain afterReadFilterChain = new FilterChain();

    private FilterChain beforeWriteFilterChain = new FilterChain();

    private FilterChain afterWriteFilterChain = new FilterChain();

    public Filterable addBeforeReadFilter(Filter filter) {
        Assert.notNull(filter);
        beforeReadFilterChain.addFilter(filter);
        return this;
    }

    public Filterable addBeforeReadFilters(Collection<Filter> filters) {
        Assert.notEmpty(filters);
        beforeReadFilterChain.addFilters(filters);
        return this;
    }

    public Filterable addAfterReadFilter(Filter filter) {
        Assert.notNull(filter);
        afterReadFilterChain.addFilter(filter);
        return this;
    }

    public Filterable addAfterReadFilters(Collection<Filter> filters) {
        Assert.notEmpty(filters);
        afterReadFilterChain.addFilters(filters);
        return this;
    }

    public Filterable addBeforeWriteFilter(Filter filter) {
        Assert.notNull(filter);
        beforeWriteFilterChain.addFilter(filter);
        return this;
    }

    public Filterable addBeforeWriteFilters(Collection<Filter> filters) {
        Assert.notEmpty(filters);
        beforeWriteFilterChain.addFilters(filters);
        return this;
    }

    public Filterable addAfterWriteFilter(Filter filter) {
        Assert.notNull(filter);
        afterWriteFilterChain.addFilter(filter);
        return this;
    }

    public Filterable addAfterWriteFilters(Collection<Filter> filters) {
        Assert.notEmpty(filters);
        afterWriteFilterChain.addFilters(filters);
        return this;
    }

    public FilterChain getBeforeReadFilterChain() {
        return beforeReadFilterChain;
    }

    public void setBeforeReadFilterChain(FilterChain beforeReadFilterChain) {
        this.beforeReadFilterChain = beforeReadFilterChain;
    }

    public FilterChain getAfterReadFilterChain() {
        return afterReadFilterChain;
    }

    public void setAfterReadFilterChain(FilterChain afterReadFilterChain) {
        this.afterReadFilterChain = afterReadFilterChain;
    }

    public FilterChain getBeforeWriteFilterChain() {
        return beforeWriteFilterChain;
    }

    public void setBeforeWriteFilterChain(FilterChain beforeWriteFilterChain) {
        this.beforeWriteFilterChain = beforeWriteFilterChain;
    }

    public FilterChain getAfterWriteFilterChain() {
        return afterWriteFilterChain;
    }

    public void setAfterWriteFilterChain(FilterChain afterWriteFilterChain) {
        this.afterWriteFilterChain = afterWriteFilterChain;
    }


}
