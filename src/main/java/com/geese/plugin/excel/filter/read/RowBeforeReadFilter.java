package com.geese.plugin.excel.filter.read;

import com.geese.plugin.excel.filter.ReadFilter;
import com.geese.plugin.excel.filter.RowFilter;
import com.geese.plugin.excel.mapping.SheetMapping;
import org.apache.poi.ss.usermodel.Row;

/**
 * 读取一行之前进行过滤
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:22
 * @sine 0.0.1
 */
public interface RowBeforeReadFilter extends RowFilter<Row, SheetMapping>, ReadFilter<Row, SheetMapping> {
}
