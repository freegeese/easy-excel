package com.geese.plugin.excel.filter.read;

import com.geese.plugin.excel.filter.RowFilter;
import com.geese.plugin.excel.mapping.SheetMapping;
import org.apache.poi.ss.usermodel.Row;

/**
 * 在读取一行之后进行过滤
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:21
 * @sine 0.0.1
 */
public interface RowAfterReadFilter extends RowFilter<Row, SheetMapping> {
}
