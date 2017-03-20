package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.FilterChain;
import com.geese.plugin.excel.mapping.CellMapping;
import com.geese.plugin.excel.mapping.SheetMapping;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.util.*;

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
            SheetMapping sheetMapping = null;
            if ("readSheet".equals(name)) {
                sheetMapping = (SheetMapping) args[1];
                filterChain = sheetMapping.getExcelMapping().getSheetAfterReadFilterChain(sheetMapping.getName());
            }
            if ("readRow".equals(name)) {
                sheetMapping = (SheetMapping) args[1];
                // 图片处理
                if (null != args[0]) {
                    readPicture((Row) args[0], sheetMapping, (Map) returnValue);
                }
                filterChain = sheetMapping.getExcelMapping().getRowAfterReadFilterChain(sheetMapping.getName());
            }
            if ("readCell".equals(name)) {
                // 图片处理
                Cell cell = (Cell) args[0];
                if (null != cell) {
                    Map<Integer, PictureData> pictureDataMap = ExcelHelper.getPictures(cell.getRow());
                    if (null != pictureDataMap && !pictureDataMap.isEmpty()) {
                        for (Map.Entry<Integer, PictureData> entry : pictureDataMap.entrySet()) {
                            if (entry.getKey() == cell.getColumnIndex()) {
                                returnValue = entry.getValue();
                                break;
                            }
                        }
                    }
                }
            }
            if (null != filterChain && !filterChain.isEmpty()) {
                if (!filterChain.doFilter(args[0], returnValue, sheetMapping, ExcelTemplate.getContext())) {
                    return ExcelOperations.EXCEL_NOT_PASS_FILTERED;
                }
            }
        }
        // 写后过滤
        if (name.startsWith("write")) {
            FilterChain filterChain = null;
            SheetMapping sheetMapping = null;
            if ("writeSheet".equals(name)) {
                sheetMapping = (SheetMapping) args[1];
                filterChain = sheetMapping.getExcelMapping().getSheetAfterWriteFilterChain(sheetMapping.getName());
            }
            if ("writeRow".equals(name)) {
                sheetMapping = (SheetMapping) args[1];
                // 图片处理
                writePicture((Row) args[0], sheetMapping, (Map<String, Object>) args[2]);
                filterChain = sheetMapping.getExcelMapping().getRowAfterWriteFilterChain(sheetMapping.getName());
            }
            // 图片处理
            if ("writeCell".equals(name)) {
                Object cellData = args[2];
                if (null != cellData && cellData instanceof byte[]) {
                    ExcelHelper.setPicture((Cell) args[0], (byte[]) cellData);
                }
            }
            if (null != filterChain && !filterChain.isEmpty()) {
                // args[2] -> target data
                if (!filterChain.doFilter(args[0], args.length == 3 ? args[2] : null, sheetMapping, ExcelTemplate.getContext())) {
                    ExcelTemplate.getContext().put(ExcelOperations.EXCEL_NOT_PASS_FILTERED, ExcelOperations.EXCEL_NOT_PASS_FILTERED);
                    return ExcelOperations.EXCEL_NOT_PASS_FILTERED;
                }
            }
        }

        return returnValue;
    }

    private void readPicture(Row row, SheetMapping sheetMapping, Map rowData) {
        Map<Integer, PictureData> collIndexAndPictureDataMap = ExcelHelper.getPictures(row);
        if (null != collIndexAndPictureDataMap && !collIndexAndPictureDataMap.isEmpty()) {
            List<CellMapping> heads = sheetMapping.getTableHeads();
            Map<Integer, String> headIndexAndHeadNameMap = new LinkedHashMap<>();
            for (CellMapping head : heads) {
                headIndexAndHeadNameMap.put(head.getColumnNumber(), head.getDataKey());
            }
            for (Map.Entry<Integer, PictureData> entry : collIndexAndPictureDataMap.entrySet()) {
                Integer columnIndex = entry.getKey();
                String columnName = headIndexAndHeadNameMap.get(columnIndex);
                rowData.put(columnName, entry.getValue().getData());
            }
        }
    }

    private void writePicture(Row row, SheetMapping sheetMapping, Map<String, Object> rowData) {
        if (null != row && !rowData.isEmpty()) {
            List<CellMapping> heads = sheetMapping.getTableHeads();
            Map<String, Integer> headNameAndHeadIndexMap = new LinkedHashMap<>();
            for (CellMapping head : heads) {
                headNameAndHeadIndexMap.put(head.getDataKey(), head.getColumnNumber());
            }
            for (Map.Entry<String, Object> entry : rowData.entrySet()) {
                Object cellData = entry.getValue();
                if (cellData instanceof byte[]) {
                    if (!headNameAndHeadIndexMap.containsKey(entry.getKey())) {
                        // 没有找到对应的列
                        continue;
                    }
                    Integer cellNumber = headNameAndHeadIndexMap.get(entry.getKey());
                    Cell cell = ExcelHelper.createCell(row, cellNumber);
                    ExcelHelper.setPicture(cell, (byte[]) cellData);
                }
            }
        }

    }
}
