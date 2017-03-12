package com.geese.plugin.excel.mapping;

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
