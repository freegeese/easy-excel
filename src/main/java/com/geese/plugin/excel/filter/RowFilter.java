package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.config.Table;
import org.apache.poi.ss.usermodel.Row;


/**
 * 行过滤器标识
 *
 * @param <T> 目标类型是Row
 * @param <M> 配置类型是Table
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:23
 * @sine 0.0.1
 */
public interface RowFilter<T extends Row, M extends Table> extends Filter<T, M> {
}
