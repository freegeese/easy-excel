package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.mapping.CellMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import com.geese.plugin.excel.mapping.SheetMapping;
import com.geese.plugin.excel.util.Assert;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Proxy;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelSimpleReader {

    private InputStream excelFrom;
    private String query;
    private String sheetName;
    private Integer sheetIndex;
    private String sheetDataKey;

    private Integer startRow;
    private Integer endRow;

    public static ExcelSimpleReader newInstance(InputStream excelFrom) {
        ExcelSimpleReader instance = new ExcelSimpleReader();
        instance.startRow = 0;
        instance.excelFrom = excelFrom;
        return instance;
    }

    public ExcelSimpleReader select(String query) {
        Assert.notEmpty(query);
        this.query = query;
        return this;
    }

    public ExcelSimpleReader from(String sheetName) {
        return from(sheetName, sheetName);
    }

    public ExcelSimpleReader from(Integer sheetIndex) {
        return from(sheetIndex, String.valueOf(sheetIndex));
    }

    public ExcelSimpleReader from(String sheetName, String sheetDataKey) {
        Assert.notEmpty(sheetName, sheetDataKey);
        this.sheetName = sheetName;
        this.sheetDataKey = sheetDataKey;
        return this;
    }

    public ExcelSimpleReader from(Integer sheetIndex, String sheetDataKey) {
        Assert.notEmpty(sheetIndex, sheetDataKey);
        this.sheetIndex = sheetIndex;
        this.sheetDataKey = sheetDataKey;
        return this;
    }

    public ExcelSimpleReader limit(Integer startRow) {
        Assert.notNull(startRow);
        this.startRow = startRow;
        return this;
    }

    public ExcelSimpleReader limit(Integer startRow, Integer maxRow) {
        Assert.notNull(startRow, maxRow);
        this.startRow = startRow;
        this.endRow = maxRow + startRow;
        return this;
    }

    public ExcelSimpleReader filter(Filter filter) {

        return this;
    }

    public Object execute() throws IOException, InvalidFormatException {
        ExcelMapping excelMapping = new ExcelMapping();

        SheetMapping sheetMapping = new SheetMapping();
        sheetMapping.setName(sheetName);
        sheetMapping.setIndex(sheetIndex);
        sheetMapping.setDataKey(sheetDataKey);
        sheetMapping.setStartRow(startRow);
        sheetMapping.setEndRow(endRow);

        List<CellMapping> tableHeads = new ArrayList<>();
        String[] columns = query.trim().split(",");
        for (String column : columns) {
            String[] indexAndName = column.trim().split("\\s+");
            int index = Integer.parseInt(indexAndName[0]);
            String name = indexAndName[1];
            CellMapping cellMapping = new CellMapping();
            cellMapping.setColumnNumber(index);
            cellMapping.setDataKey(name);
            tableHeads.add(cellMapping);
        }
        sheetMapping.setTableHeads(tableHeads);
        excelMapping.setSheetMappings(Arrays.asList(sheetMapping));
        Workbook workbook = WorkbookFactory.create(excelFrom);
        ExcelTemplate excelTemplate = new ExcelTemplate();

        ExcelOperationsProxy handler = new ExcelOperationsProxy(excelTemplate);
        Class<? extends ExcelTemplate> targetClass = excelTemplate.getClass();
        ClassLoader loader = targetClass.getClassLoader();
        Class<?>[] interfaces = targetClass.getInterfaces();
        ExcelOperations proxy = (ExcelOperations) Proxy.newProxyInstance(loader, interfaces, handler);

        Object result = proxy.readExcel(workbook, excelMapping);

        return result;
    }

}
