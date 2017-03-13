package com.geese.plugin.excel.mapping;

import com.geese.plugin.excel.ExcelHelper;
import com.geese.plugin.excel.OperationKey;
import com.geese.plugin.excel.filter.Filter;

import java.io.InputStream;
import java.util.*;

/**
 * Created by Administrator on 2017/3/12.
 */
public class ClientMapping {

    private InputStream excelInput;

    private List<String> queries = new ArrayList<>();

    private Map<String, Collection<Filter>> sheetFiltersMap = new LinkedHashMap<>();

    public ClientMapping addFilter(Filter filter, String switchSheet) {
        if (sheetFiltersMap.containsKey(switchSheet)) {
            sheetFiltersMap.get(switchSheet).add(filter);
            return this;
        }
        Set<Filter> filters = new LinkedHashSet<>();
        filters.add(filter);
        sheetFiltersMap.put(switchSheet, filters);
        return this;
    }

    public ClientMapping addFilters(Collection<Filter> filters, String switchSheet) {
        if (sheetFiltersMap.containsKey(switchSheet)) {
            sheetFiltersMap.get(switchSheet).addAll(filters);
            return this;
        }
        sheetFiltersMap.put(switchSheet, filters);
        return this;
    }

    /**
     * 解析客户端输入
     *
     * @return
     */
    public ExcelMapping parseClientInput() {
        ExcelMapping excelMapping = new ExcelMapping();
        Map<String, SheetMapping> sheetMappingMap = new LinkedHashMap<>();
        // 解析查询语句
        for (String query : queries) {
            query = query.trim();
            // 匹配判断是否是散列点查询
            if (query.matches("^\\{.+\\}$")) {
                query = query.replaceAll("\\{|\\}", "");
                // 关键字与查询语句映射
                Map<OperationKey, String> operationKeyMap = ExcelHelper.parseQuery(query);
                // 如果已经创建 sheet mapping 则获取已创建好的
                String from = operationKeyMap.get(OperationKey.FROM);
                SheetMapping sheetMapping = sheetMappingMap.get(from);
                if (null == sheetMapping) {
                    sheetMapping = new SheetMapping();
                    sheetMapping.setDataKey(from);
                    sheetMapping.setName(from);
                    sheetMappingMap.put(from, sheetMapping);
                    sheetMapping.setExcelMapping(excelMapping);
                }

                // 格式：1-3 name, 2-4 age
                String[] points = operationKeyMap.get(OperationKey.COLUMN).split(",");
                for (String point : points) {
                    String[] items = point.trim().split("-|\\s+");
                    CellMapping p = new CellMapping();
                    p.setRowNumber(Integer.valueOf(items[0]));
                    p.setColumnNumber(Integer.valueOf(items[1]));
                    p.setDataKey(items[2]);
                    sheetMapping.getPoints().add(p);
                    p.setSheetMapping(sheetMapping);
                }
                continue;
            }

            // 列表查询
            // 关键字 与 查询语句映射
            Map<OperationKey, String> operationKeyMap = ExcelHelper.parseQuery(query);
            SheetMapping sheetMapping = new SheetMapping();
            sheetMapping.setExcelMapping(excelMapping);
            // 列索引 和 列名称 (select 1 name, 2 age)
            String[] columns = operationKeyMap.get(OperationKey.COLUMN).split(",");
            // 自动计算的列索引 column index
            int autoColumnIndex = 0;
            for (String column : columns) {
                // 表格头部列  table head column
                CellMapping cellMapping = new CellMapping();
                String[] indexAndName = column.trim().split("\\s+");
                if (indexAndName.length == 1) {
                    cellMapping.setDataKey(indexAndName[0].trim());
                    cellMapping.setColumnNumber(autoColumnIndex++);
                } else {
                    Integer columnIndex = Integer.valueOf(indexAndName[0].trim());
                    cellMapping.setColumnNumber(columnIndex);
                    cellMapping.setDataKey(indexAndName[1].trim());
                    autoColumnIndex = columnIndex + 1;
                }
                // 添加头部列
                sheetMapping.getTableHeads().add(cellMapping);
                // 关联到Sheet Mapping
                cellMapping.setSheetMapping(sheetMapping);
            }
            // 查询结果来自哪个Sheet (from Sheet1)
            String from = operationKeyMap.get(OperationKey.FROM);
            sheetMapping.setName(from);
            sheetMapping.setDataKey(from);

            // 分页查询 (limit 10,10)
            String limit = operationKeyMap.get(OperationKey.LIMIT);
            if (null != limit) {
                String[] startAndSize = limit.trim().split(",");
                sheetMapping.setStartRow(Integer.valueOf(startAndSize[0].trim()));
                if (startAndSize.length == 2) {
                    Integer size = Integer.valueOf(startAndSize[1].trim());
                    sheetMapping.setEndRow(sheetMapping.getStartRow() + size);
                }
            }
            sheetMappingMap.put(from, sheetMapping);

        }

        // 设置关联的sheet
        excelMapping.setSheetMappings(sheetMappingMap.values());
        // 对过滤器进行分类
        excelMapping.setSheetFiltersMap(getSheetFiltersMap());
        excelMapping.classificationFilters();
        return excelMapping;
    }

    public InputStream getExcelInput() {
        return excelInput;
    }

    public void setExcelInput(InputStream excelInput) {
        this.excelInput = excelInput;
    }

    public List<String> getQueries() {
        return queries;
    }

    public void setQueries(List<String> queries) {
        this.queries = queries;
    }

    public Map<String, Collection<Filter>> getSheetFiltersMap() {
        return sheetFiltersMap;
    }

    public void setSheetFiltersMap(Map<String, Collection<Filter>> sheetFiltersMap) {
        this.sheetFiltersMap = sheetFiltersMap;
    }


}
