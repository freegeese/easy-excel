package com.geese.plugin.excel.filter;

import java.util.Collection;

/**
 * <p> 可过滤的 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:00
 * @sine 0.0.2
 */
public abstract class Filterable {

    private FilterChain beforeReadFilterChain;

    private FilterChain afterReadFilterChain;

    private FilterChain beforeWriteFilterChain;

    private FilterChain afterWriteFilterChain;

    public Filterable addBeforeReadFilter(Filter filter) {
        if (null == beforeReadFilterChain) {
            beforeReadFilterChain = new FilterChain();
        }
        beforeReadFilterChain.addFilter(filter);
        return this;
    }

    public Filterable addBeforeReadFilters(Collection<Filter> filters) {
        if (null == beforeReadFilterChain) {
            beforeReadFilterChain = new FilterChain();
        }
        beforeReadFilterChain.addFilters(filters);
        return this;
    }

    public Filterable addAfterReadFilter(Filter filter) {
        if (null == afterReadFilterChain) {
            afterReadFilterChain = new FilterChain();
        }
        afterReadFilterChain.addFilter(filter);
        return this;
    }

    public Filterable addAfterReadFilters(Collection<Filter> filters) {
        if (null == afterReadFilterChain) {
            afterReadFilterChain = new FilterChain();
        }
        afterReadFilterChain.addFilters(filters);
        return this;
    }

    public Filterable addBeforeWriteFilter(Filter filter) {
        if (null == beforeWriteFilterChain) {
            beforeWriteFilterChain = new FilterChain();
        }
        beforeWriteFilterChain.addFilter(filter);
        return this;
    }

    public Filterable addBeforeWriteFilters(Collection<Filter> filters) {
        if (null == beforeWriteFilterChain) {
            beforeWriteFilterChain = new FilterChain();
        }
        beforeWriteFilterChain.addFilters(filters);
        return this;
    }

    public Filterable addAfterWriteFilter(Filter filter) {
        if (null == afterWriteFilterChain) {
            afterWriteFilterChain = new FilterChain();
        }
        afterWriteFilterChain.addFilter(filter);
        return this;
    }

    public Filterable addAfterWriteFilters(Collection<Filter> filters) {
        if (null == afterWriteFilterChain) {
            afterWriteFilterChain = new FilterChain();
        }
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
