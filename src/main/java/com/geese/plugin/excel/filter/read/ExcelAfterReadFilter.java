package com.geese.plugin.excel.filter.read;

import com.geese.plugin.excel.filter.ExcelFilter;
import com.geese.plugin.excel.filter.ReadFilter;
import com.geese.plugin.excel.mapping.CellMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p> 读取excel后过滤 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:19
 * @sine 0.0.2
 */
public interface ExcelAfterReadFilter extends ExcelFilter<Workbook, ExcelMapping>, ReadFilter<Workbook, ExcelMapping> {
}
