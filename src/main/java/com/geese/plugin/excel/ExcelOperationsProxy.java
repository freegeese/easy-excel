package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.FilterChain;
import com.geese.plugin.excel.mapping.SheetMapping;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelOperationsProxy implements InvocationHandler {
    // 目标对象
    private ExcelOperations target;

    // 构建代理对象的时候织入目标对象
    public ExcelOperationsProxy(ExcelOperations target) {
        this.target = target;
    }

    @Override
    public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
        String name = method.getName();
        if (name.startsWith("read")) {
            FilterChain readBeforeFilterChain = null;
            Object mapping = null;
            if ("readSheet".equals(name)) {
                SheetMapping sheetMapping = (SheetMapping) args[1];
                mapping = args[1];
                readBeforeFilterChain = sheetMapping.getExcelMapping().getSheetBeforeReadFilterChain(sheetMapping.getName());
            }
            if ("readRow".equals(name)) {
                SheetMapping sheetMapping = (SheetMapping) args[1];
                mapping = args[1];
                readBeforeFilterChain = sheetMapping.getExcelMapping().getRowBeforeReadFilterChain(sheetMapping.getName());
            }
            if (null != readBeforeFilterChain && !readBeforeFilterChain.isEmpty()) {
                if (!readBeforeFilterChain.doFilter(args[0], null, mapping)) {
                    return ExcelOperations.EXCEL_NOT_FILTERED;
                }
            }
        }

        Object returnValue = method.invoke(target, args);

        if (name.startsWith("read")) {
            FilterChain readAfterFilterChain = null;
            Object mapping = null;
            if ("readSheet".equals(name)) {
                SheetMapping sheetMapping = (SheetMapping) args[1];
                mapping = args[1];
                readAfterFilterChain = sheetMapping.getExcelMapping().getSheetAfterReadFilterChain(sheetMapping.getName());
            }
            if ("readRow".equals(name)) {
                SheetMapping sheetMapping = (SheetMapping) args[1];
                mapping = args[1];
                readAfterFilterChain = sheetMapping.getExcelMapping().getRowAfterReadFilterChain(sheetMapping.getName());
            }
            if (null != readAfterFilterChain && !readAfterFilterChain.isEmpty()) {
                if (!readAfterFilterChain.doFilter(args[0], returnValue, mapping)) {
                    return ExcelOperations.EXCEL_NOT_FILTERED;
                }
            }
        }

        return returnValue;
    }
}
