package com.geese.plugin.excel.filter.write;

import com.geese.plugin.excel.filter.RowFilter;
import com.geese.plugin.excel.filter.WriteFilter;
import com.geese.plugin.excel.mapping.SheetMapping;
import org.apache.poi.ss.usermodel.Row;

/**
 * 在写入一行之前进行过滤
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:25
 * @sine 0.0.1
 */
public interface RowBeforeWriteFilter extends RowFilter<Row, SheetMapping>, WriteFilter<Row, SheetMapping> {
}
