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
        // 读前过滤
        if (name.startsWith("read")) {
            FilterChain filterChain = null;
            Object mapping = null;
            if ("readSheet".equals(name)) {
                SheetMapping sheetMapping = (SheetMapping) args[1];
                mapping = args[1];
                filterChain = sheetMapping.getExcelMapping().getSheetBeforeReadFilterChain(sheetMapping.getName());
            }
            if ("readRow".equals(name)) {
                SheetMapping sheetMapping = (SheetMapping) args[1];
                mapping = args[1];
                filterChain = sheetMapping.getExcelMapping().getRowBeforeReadFilterChain(sheetMapping.getName());
            }
            if (null != filterChain && !filterChain.isEmpty()) {
                if (!filterChain.doFilter(args[0], null, mapping, ExcelTemplate.getContext())) {
                    return ExcelOperations.EXCEL_NOT_PASS_FILTERED;
                }
            }
        }
        // 写前过滤
        if (name.startsWith("write")) {
            FilterChain filterChain = null;
            SheetMapping mapping = null;
            if ("writeSheet".equals(name)) {
                // sheet data -> mapping.getTableData() , mapping.getPoints()
                mapping = (SheetMapping) args[1];
                filterChain = mapping.getExcelMapping().getSheetBeforeWriteFilterChain(mapping.getName());
            }
            if ("writeRow".equals(name)) {
                // row data -> args[2]
                mapping = (SheetMapping) args[1];
                filterChain = mapping.getExcelMapping().getRowBeforeWriteFilterChain(mapping.getName());
            }
            // args[2] -> target data
            if (null != filterChain && !filterChain.isEmpty()) {
                if (!filterChain.doFilter(args[0], args.length == 3 ? args[2] : null, mapping, ExcelTemplate.getContext())) {
                    ExcelTemplate.getContext().put(ExcelOperations.EXCEL_NOT_PASS_FILTERED, ExcelOperations.EXCEL_NOT_PASS_FILTERED);
                    return ExcelOperations.EXCEL_NOT_PASS_FILTERED;
                }
            }
        }

        Object returnValue = method.invoke(target, args);

        // 读后过滤
        if (name.startsWith("read")) {
            FilterChain filterChain = null;
            Object mapping = null;
            if ("readSheet".equals(name)) {
                SheetMapping sheetMapping = (SheetMapping) args[1];
                mapping = args[1];
                filterChain = sheetMapping.getExcelMapping().getSheetAfterReadFilterChain(sheetMapping.getName());
            }
            if ("readRow".equals(name)) {
                SheetMapping sheetMapping = (SheetMapping) args[1];
                mapping = args[1];
                filterChain = sheetMapping.getExcelMapping().getRowAfterReadFilterChain(sheetMapping.getName());
            }
            if (null != filterChain && !filterChain.isEmpty()) {
                if (!filterChain.doFilter(args[0], returnValue, mapping, ExcelTemplate.getContext())) {
                    return ExcelOperations.EXCEL_NOT_PASS_FILTERED;
                }
            }
        }
        // 写后过滤
        if (name.startsWith("write")) {
            FilterChain filterChain = null;
            SheetMapping mapping = null;
            if ("writeSheet".equals(name)) {
                mapping = (SheetMapping) args[1];
                filterChain = mapping.getExcelMapping().getSheetAfterWriteFilterChain(mapping.getName());
            }
            if ("writeRow".equals(name)) {
                mapping = (SheetMapping) args[1];
                filterChain = mapping.getExcelMapping().getRowAfterWriteFilterChain(mapping.getName());
            }
            if (null != filterChain && !filterChain.isEmpty()) {
                // args[2] -> target data
                if (!filterChain.doFilter(args[0], args.length == 3 ? args[2] : null, mapping, ExcelTemplate.getContext())) {
                    ExcelTemplate.getContext().put(ExcelOperations.EXCEL_NOT_PASS_FILTERED, ExcelOperations.EXCEL_NOT_PASS_FILTERED);
                    return ExcelOperations.EXCEL_NOT_PASS_FILTERED;
                }
            }
        }

        return returnValue;
    }
}
