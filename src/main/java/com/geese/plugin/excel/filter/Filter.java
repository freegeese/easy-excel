package com.geese.plugin.excel.filter;

import java.util.Map;

/**
 * ExcelMapping 过滤器接口定义
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:19
 * @sine 0.0.1
 */
public interface Filter<T, M> {

    /**
     * 过滤行或者列
     *
     * @param target  row 或者 column
     * @param data    当过滤 write 操作的时候传入的数据，如果是read过滤，data为null
     * @param mapping 当过滤 row 时，可以拿到 row 所在的 sheet 配置信息, 当过滤 cell 时，可以拿到 cell 的配置信息
     * @param context 一次读/写操作的上下文环境
     */
    boolean doFilter(T target, Object data, M mapping, Map context);

}
