package com.geese.plugin.excel.filter;

import java.util.Collection;
import java.util.LinkedList;
import java.util.List;

/**
 * 过滤链，可以添加多个过滤器
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:20
 * @sine 0.0.1
 */
public class FilterChain<T, M> {
    private List<Filter> filterList = new LinkedList<>();

    /**
     * 在链条上添加过滤器节点
     *
     * @param filter
     * @return
     */
    public FilterChain addFilter(Filter filter) {
        filterList.add(filter);
        return this;
    }

    public FilterChain addFilters(Collection<Filter> filters) {
        filterList.addAll(filters);
        return this;
    }

    /**
     * 执行过滤，会调用链条上每个节点的doFilter方法
     *
     * @param target
     * @param data
     * @param config
     */
    public boolean doFilter(T target, Object data, M config) {
        if (!filterList.isEmpty()) {
            for (Filter filter : filterList) {
                if (!filter.doFilter(target, data, config)) {
                    return false;
                }
            }
        }
        return true;
    }
}
