package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.filter.ReadFilter;
import com.geese.plugin.excel.mapping.ClientMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import com.geese.plugin.excel.util.Assert;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * 读取Excel
 */
public class ExcelReader {
    // 处理客户端输入信息
    private ClientMapping clientMapping = new ClientMapping();

    public static ExcelReader newInstance(InputStream excelInput) {
        ExcelReader instance = new ExcelReader();
        instance.clientMapping.setExcelInput(excelInput);
        return instance;
    }

    public ExcelReader select(String query) {
        Assert.notEmpty(query);
        clientMapping.getQueries().add(query);
        return this;
    }

    public ExcelReader select(String first, String second, String... more) {
        Assert.notEmpty(first, second);
        List<String> queries = clientMapping.getQueries();
        queries.add(first);
        queries.add(second);
        if (null != more && more.length > 0) {
            queries.addAll(Arrays.asList(more));
        }
        return this;
    }

    public ExcelReader filter(ReadFilter filter, String switchSheet) {
        clientMapping.addFilter(filter, switchSheet);
        return this;
    }

    public ExcelReader filters(ReadFilter[] filters, String switchSheet) {
        return filters(Arrays.asList(filters), switchSheet);
    }

    public ExcelReader filters(Collection<ReadFilter> filters, String switchSheet) {
        for (ReadFilter filter : filters) {
            clientMapping.addFilter(filter, switchSheet);
        }
        return this;
    }

    public ExcelResult execute() throws IOException, InvalidFormatException {
        // 把客户输入转换为Excel映射信息
        ExcelMapping excelMapping = clientMapping.parseClientInput();
        // 创建Workbook
        Workbook workbook = WorkbookFactory.create(clientMapping.getExcelInput());
        // Excel操作接口代理
        ExcelOperations proxy = ExcelOperationsProxyFactory.getProxy();
        Map data = (Map) proxy.readExcel(workbook, excelMapping);
        ExcelResult excelResult = new ExcelResult();
        excelResult.setData(data);
        excelResult.setContext(ExcelTemplate.getContext());
        return excelResult;
    }
}
