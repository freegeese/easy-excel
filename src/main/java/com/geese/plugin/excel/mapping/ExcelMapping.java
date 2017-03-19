package com.geese.plugin.excel.mapping;

import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.filter.FilterChain;
import com.geese.plugin.excel.filter.Filterable;
import com.geese.plugin.excel.filter.read.RowAfterReadFilter;
import com.geese.plugin.excel.filter.read.RowBeforeReadFilter;
import com.geese.plugin.excel.filter.read.SheetAfterReadFilter;
import com.geese.plugin.excel.filter.read.SheetBeforeReadFilter;
import com.geese.plugin.excel.filter.write.RowAfterWriteFilter;
import com.geese.plugin.excel.filter.write.RowBeforeWriteFilter;
import com.geese.plugin.excel.filter.write.SheetAfterWriteFilter;
import com.geese.plugin.excel.filter.write.SheetBeforeWriteFilter;

import java.io.File;
import java.util.*;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelMapping {

    // Sheet 表格映射信息
    private Collection<SheetMapping> sheetMappings;
    // 过滤器
    private Map<String, Collection<Filter>> sheetFiltersMap = new LinkedHashMap<>();

    public Collection<SheetMapping> getSheetMappings() {
        return sheetMappings;
    }

    public void setSheetMappings(Collection<SheetMapping> sheetMappings) {
        this.sheetMappings = sheetMappings;
    }

    public Map<String, Collection<Filter>> getSheetFiltersMap() {
        return sheetFiltersMap;
    }

    public void setSheetFiltersMap(Map<String, Collection<Filter>> sheetFiltersMap) {
        this.sheetFiltersMap = sheetFiltersMap;
    }

    private Map<String, FilterChain> sheetBeforeReadFilterChainMap = new HashMap<>();
    private Map<String, FilterChain> sheetBeforeWriteFilterChainMap = new HashMap<>();
    private Map<String, FilterChain> sheetAfterReadFilterChainMap = new HashMap<>();
    private Map<String, FilterChain> sheetAfterWriteFilterChainMap = new HashMap<>();

    private Map<String, FilterChain> rowBeforeReadFilterChainMap = new HashMap<>();
    private Map<String, FilterChain> rowBeforeWriteFilterChainMap = new HashMap<>();
    private Map<String, FilterChain> rowAfterReadFilterChainMap = new HashMap<>();
    private Map<String, FilterChain> rowAfterWriteFilterChainMap = new HashMap<>();

    public FilterChain getSheetBeforeReadFilterChain(String sheet) {
        return sheetBeforeReadFilterChainMap.get(sheet);
    }

    public FilterChain getSheetBeforeWriteFilterChain(String sheet) {
        return sheetBeforeWriteFilterChainMap.get(sheet);
    }

    public FilterChain getSheetAfterReadFilterChain(String sheet) {
        return sheetAfterReadFilterChainMap.get(sheet);
    }

    public FilterChain getSheetAfterWriteFilterChain(String sheet) {
        return sheetAfterWriteFilterChainMap.get(sheet);
    }

    public FilterChain getRowBeforeReadFilterChain(String sheet) {
        return rowBeforeReadFilterChainMap.get(sheet);
    }

    public FilterChain getRowBeforeWriteFilterChain(String sheet) {
        return rowBeforeWriteFilterChainMap.get(sheet);
    }

    public FilterChain getRowAfterReadFilterChain(String sheet) {
        return rowAfterReadFilterChainMap.get(sheet);
    }

    public FilterChain getRowAfterWriteFilterChain(String sheet) {
        return rowAfterWriteFilterChainMap.get(sheet);
    }

    /**
     * 过滤器分类
     */
    public void classificationFilters() {
        for (Map.Entry<String, Collection<Filter>> entry : sheetFiltersMap.entrySet()) {
            String sheet = entry.getKey();
            Collection<Filter> filters = entry.getValue();
            for (Filter filter : filters) {
                if (filter instanceof SheetBeforeReadFilter) {
                    addFilterToFilterChain(filter, sheetBeforeReadFilterChainMap, sheet);
                    continue;
                }
                if (filter instanceof SheetBeforeWriteFilter) {
                    addFilterToFilterChain(filter, sheetBeforeWriteFilterChainMap, sheet);
                    continue;
                }
                if (filter instanceof SheetAfterReadFilter) {
                    addFilterToFilterChain(filter, sheetAfterReadFilterChainMap, sheet);
                    continue;
                }
                if (filter instanceof SheetAfterWriteFilter) {
                    addFilterToFilterChain(filter, sheetAfterWriteFilterChainMap, sheet);
                    continue;
                }

                if (filter instanceof RowBeforeReadFilter) {
                    addFilterToFilterChain(filter, rowBeforeReadFilterChainMap, sheet);
                    continue;
                }
                if (filter instanceof RowBeforeWriteFilter) {
                    addFilterToFilterChain(filter, rowBeforeWriteFilterChainMap, sheet);
                    continue;
                }
                if (filter instanceof RowAfterReadFilter) {
                    addFilterToFilterChain(filter, rowAfterReadFilterChainMap, sheet);
                    continue;
                }
                if (filter instanceof RowAfterWriteFilter) {
                    addFilterToFilterChain(filter, rowAfterWriteFilterChainMap, sheet);
                    continue;
                }
            }
        }
    }

    private void addFilterToFilterChain(Filter filter, Map<String, FilterChain> filterChainMap, String key) {
        if (filterChainMap.containsKey(key)) {
            filterChainMap.get(key).addFilter(filter);
            return;
        }
        FilterChain filterChain = new FilterChain();
        filterChain.addFilter(filter);
        filterChainMap.put(key, filterChain);
    }

}
