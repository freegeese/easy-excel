package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.config.Table;
import org.apache.poi.ss.usermodel.Row;

/**
 * <p> 写入一行之后过滤 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:14
 * @sine 0.0.2
 */
public interface RowAfterWriteFilter extends RowFilter<Row, Table> {
}
