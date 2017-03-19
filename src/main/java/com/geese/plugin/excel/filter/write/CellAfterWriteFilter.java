package com.geese.plugin.excel.filter.write;

import com.geese.plugin.excel.filter.CellFilter;
import com.geese.plugin.excel.filter.WriteFilter;
import com.geese.plugin.excel.mapping.CellMapping;
import org.apache.poi.ss.usermodel.Cell;

/**
 * <p> 写入单元格后过滤 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:13
 * @sine 0.0.2
 */
public interface CellAfterWriteFilter extends CellFilter<Cell, CellMapping>, WriteFilter<Cell, CellMapping> {
}
