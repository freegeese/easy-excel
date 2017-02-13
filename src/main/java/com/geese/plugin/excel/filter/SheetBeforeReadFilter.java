package com.geese.plugin.excel.filter;

import com.geese.plugin.excel.config.MySheet;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * <p> 读取sheet之前过滤 <br>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/21 22:15
 * @sine 0.0.2
 */
public interface SheetBeforeReadFilter extends SheetFilter<Sheet, MySheet> {
}
