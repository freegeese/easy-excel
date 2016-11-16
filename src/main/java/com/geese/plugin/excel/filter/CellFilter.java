package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.config.Point;
import org.apache.poi.ss.usermodel.Cell;

/**
 * 列过滤器标识
 *
 * @param <T>
 * @param <M>
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:26
 * @sine 0.0.1
 */
public interface CellFilter<T extends Cell, M extends Point> extends Filter<T, M> {
}
