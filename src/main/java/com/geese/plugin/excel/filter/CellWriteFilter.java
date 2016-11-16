package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.config.Point;
import org.apache.poi.ss.usermodel.Cell;

/**
 * 写入一列之前过滤
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:27
 * @sine 0.0.1
 */
public interface CellWriteFilter extends CellFilter<Cell, Point> {
}
