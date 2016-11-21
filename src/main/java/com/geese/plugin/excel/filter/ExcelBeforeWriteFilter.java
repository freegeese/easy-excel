package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.config.Excel;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p> 写入excel前过滤 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:19
 * @sine 0.0.2
 */
public interface ExcelBeforeWriteFilter extends ExcelFilter<Workbook, Excel> {
}
