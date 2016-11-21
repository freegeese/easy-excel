package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.config.Excel;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p> excel 过滤器 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:09
 * @sine 0.0.2
 */
public interface ExcelFilter<T extends Workbook, M extends Excel> extends Filter<T, M> {
}
