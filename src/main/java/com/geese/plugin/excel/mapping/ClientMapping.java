package com.geese.plugin.excel.mapping;

import com.geese.plugin.excel.ExcelHelper;
import com.geese.plugin.excel.OperationKey;
import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.filter.ReadFilter;
import com.geese.plugin.excel.util.Assert;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

/**
 * Created by Administrator on 2017/3/12.
 */
public class ClientMapping {

    private InputStream excelInput;
    private List<String> queries = new ArrayList<>();

    private OutputStream excelOutput;
    private List<String> inserts = new ArrayList<>();
    private Boolean useXlsFormat = true;
    private File excelOutputTemplate;
    private Map<String, List<Map>> sheetAndTableDataMap = new LinkedHashMap<>();
    private Map<String, Map> sheetAndPointDataMap = new LinkedHashMap<>();

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

    public ClientMapping addTableData(List<Map> tableData, String switchSheet) {
        sheetAndTableDataMap.put(switchSheet, tableData);
        return this;
    }

    public ClientMapping addPointData(Map pointData, String switchSheet) {
        sheetAndPointDataMap.put(switchSheet, pointData);
        return this;
    }

    /**
     * 解析客户端输入
     *
     * @return
     */
    public ExcelMapping parseClientInput() {
        ExcelMapping excelMapping = null;
        // 解析查询语句
        if (!queries.isEmpty()) {
            excelMapping = parseQuery();
        } else if (!inserts.isEmpty()) {
            excelMapping = parseInsert();
        }
        // 对过滤器进行分类
        excelMapping.setSheetFiltersMap(getSheetFiltersMap());
        excelMapping.classificationFilters();
        return excelMapping;
    }

    /**
     * 解析插入语句
     *
     * @return
     */
    private ExcelMapping parseInsert() {
        ExcelMapping excelMapping = new ExcelMapping();
        Map<String, SheetMapping> sheetMappingMap = new LinkedHashMap<>();
        for (String insert : inserts) {
            insert = insert.trim();
            // 匹配判断是否是散列点查询
            if (insert.matches("^\\{.+\\}$")) {
                insert = insert.replaceAll("\\{|\\}", "");
                // 关键字与查询语句映射
                Map<OperationKey, String> operationKeyMap = ExcelHelper.parseInsert(insert);
                // 如果已经创建 sheet mapping 则获取已创建好的
                String into = operationKeyMap.get(OperationKey.INTO);
                // 确定有数据
                Assert.isTrue(sheetAndPointDataMap.containsKey(into));

                SheetMapping sheetMapping = sheetMappingMap.get(into);
                if (null == sheetMapping) {
                    sheetMapping = new SheetMapping();
                    sheetMapping.setDataKey(into);
                    sheetMapping.setName(into);
                    if (ExcelHelper.isNumber(into)) {
                        sheetMapping.setIndex(Integer.valueOf(into));
                    }
                    sheetMappingMap.put(into, sheetMapping);
                    sheetMapping.setExcelMapping(excelMapping);
                }

                // 格式：1-3 name, 2-4 age
                Map pointData = sheetAndPointDataMap.get(into);
                String[] points = operationKeyMap.get(OperationKey.COLUMN).split(",");
                for (String point : points) {
                    String[] items = point.trim().split("-|\\s+");
                    CellMapping p = new CellMapping();
                    p.setRowNumber(Integer.valueOf(items[0]));
                    p.setColumnNumber(Integer.valueOf(items[1]));
                    p.setDataKey(items[2]);
                    p.setData(pointData.get(p.getDataKey()));
                    sheetMapping.getPoints().add(p);
                    p.setSheetMapping(sheetMapping);
                }
                continue;
            }

            // 列表插入
            // 关键字与插入语句映射
            Map<OperationKey, String> operationKeyMap = ExcelHelper.parseInsert(insert);
            String into = operationKeyMap.get(OperationKey.INTO);
            // 确定有数据
            Assert.isTrue(sheetAndTableDataMap.containsKey(into));

            SheetMapping sheetMapping = new SheetMapping();
            sheetMapping.setExcelMapping(excelMapping);
            // 列索引 和 列名称 (insert 1 name, 2 age)
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
            // 数据插入到哪个Sheet (into Sheet1)
            sheetMapping.setName(into);
            if (ExcelHelper.isNumber(into)) {
                sheetMapping.setIndex(Integer.valueOf(into));
            }
            sheetMapping.setDataKey(into);

            // 分页查询 (limit 10,10)
            String limit = operationKeyMap.get(OperationKey.LIMIT);
            if (null != limit) {
                String[] items = limit.trim().split(",");
                // 开始行
                sheetMapping.setStartRow(Integer.valueOf(items[0].trim()));
                if (items.length == 2) {
                    // 行间隔
                    sheetMapping.setRowInterval(Integer.valueOf(items[1].trim()));
                }
            }
            sheetMapping.setTableData(sheetAndTableDataMap.get(into));
            sheetMappingMap.put(into, sheetMapping);
        }
        excelMapping.setSheetMappings(sheetMappingMap.values());
        return excelMapping;
    }

    /**
     * 解析查询语句
     *
     * @return
     */
    private ExcelMapping parseQuery() {
        ExcelMapping excelMapping = new ExcelMapping();
        Map<String, SheetMapping> sheetMappingMap = new LinkedHashMap<>();
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
                    if (ExcelHelper.isNumber(from)) {
                        sheetMapping.setIndex(Integer.valueOf(from));
                    }
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
            if (ExcelHelper.isNumber(from)) {
                sheetMapping.setIndex(Integer.valueOf(from));
            }
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
        excelMapping.setSheetMappings(sheetMappingMap.values());
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

    public List<String> getInserts() {
        return inserts;
    }

    public void setInserts(List<String> inserts) {
        this.inserts = inserts;
    }

    public OutputStream getExcelOutput() {
        return excelOutput;
    }

    public void setExcelOutput(OutputStream excelOutput) {
        this.excelOutput = excelOutput;
    }

    public Boolean getUseXlsFormat() {
        return useXlsFormat;
    }

    public void setUseXlsFormat(Boolean useXlsFormat) {
        this.useXlsFormat = useXlsFormat;
    }

    public File getExcelOutputTemplate() {
        return excelOutputTemplate;
    }

    public void setExcelOutputTemplate(File excelOutputTemplate) {
        this.excelOutputTemplate = excelOutputTemplate;
    }

    public Map<String, Collection<Filter>> getSheetFiltersMap() {
        return sheetFiltersMap;
    }

    public void setSheetFiltersMap(Map<String, Collection<Filter>> sheetFiltersMap) {
        this.sheetFiltersMap = sheetFiltersMap;
    }


}
