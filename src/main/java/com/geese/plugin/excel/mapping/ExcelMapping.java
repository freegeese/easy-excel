package com.geese.plugin.excel.mapping;

import java.io.File;
import java.util.List;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelMapping {

    // 模板(用于写入数据)
    private File template;
    // Sheet 表格映射信息
    private List<SheetMapping> sheetMappings;

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
}
