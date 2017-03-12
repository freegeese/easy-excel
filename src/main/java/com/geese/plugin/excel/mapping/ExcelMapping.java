package com.geese.plugin.excel.mapping;

import com.geese.plugin.excel.filter.Filter;

import java.io.File;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelMapping {

    // 模板(用于写入数据)
    private File template;
    // Sheet 表格映射信息
    private List<SheetMapping> sheetMappings;
    // 过滤器
    private Map<String, Collection<Filter>> sheetFiltersMap = new LinkedHashMap<>();

    public File getTemplate() {
        return template;
    }

    public void setTemplate(File template) {
        this.template = template;
    }

    public List<SheetMapping> getSheetMappings() {
        return sheetMappings;
    }

    public void setSheetMappings(List<SheetMapping> sheetMappings) {
        this.sheetMappings = sheetMappings;
    }

    public Map<String, Collection<Filter>> getSheetFiltersMap() {
        return sheetFiltersMap;
    }

    public void setSheetFiltersMap(Map<String, Collection<Filter>> sheetFiltersMap) {
        this.sheetFiltersMap = sheetFiltersMap;
    }
}
