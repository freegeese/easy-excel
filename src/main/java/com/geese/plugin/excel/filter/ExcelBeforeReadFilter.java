package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.config.Excel;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p> 读取excel前过滤 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:18
 * @sine 0.0.2
 */
public interface ExcelBeforeReadFilter extends ExcelFilter<Workbook, Excel> {
}
