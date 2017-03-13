package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.mapping.ClientMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import com.geese.plugin.excel.util.Assert;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Proxy;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2017/3/12.
 */
public class ExcelStandardReader {
    // 处理客户端输入信息
    private ClientMapping clientMapping = new ClientMapping();

    public static ExcelStandardReader newInstance(InputStream excelInput) {
        ExcelStandardReader instance = new ExcelStandardReader();
        instance.getClientMapping().setExcelInput(excelInput);
        return instance;
    }

    public ExcelStandardReader select(String query) {
        Assert.notEmpty(query);
        clientMapping.getQueries().add(query);
        return this;
    }

    public ExcelStandardReader select(String first, String second, String... more) {
        Assert.notEmpty(first, second);
        List<String> queries = clientMapping.getQueries();
        queries.add(first);
        queries.add(second);
        if (null != more && more.length > 0) {
            queries.addAll(Arrays.asList(more));
        }
        return this;
    }

    public ExcelStandardReader filter(Filter filter, String switchSheet) {
        getClientMapping().addFilter(filter, switchSheet);
        return this;
    }

    public ExcelStandardReader filters(List<Filter> filters, String switchSheet) {
        getClientMapping().addFilters(filters, switchSheet);
        return this;
    }

    public Map execute() throws IOException, InvalidFormatException {
        // 客户输入
        ClientMapping clientMapping = getClientMapping();
        // 把客户输入转换为Excel映射信息
        ExcelMapping excelMapping = clientMapping.parseClientInput();
        // 创建Workbook
        Workbook workbook = WorkbookFactory.create(clientMapping.getExcelInput());
        // 创建Excel模板
        ExcelTemplate excelTemplate = new ExcelTemplate();
        // Excel操作接口代理
        ExcelOperations proxy = ExcelOperationsProxyFactory.getProxy();
        return (Map) proxy.readExcel(workbook, excelMapping);
    }

    public ClientMapping getClientMapping() {
        return clientMapping;
    }

    public void setClientMapping(ClientMapping clientMapping) {
        this.clientMapping = clientMapping;
    }
}
