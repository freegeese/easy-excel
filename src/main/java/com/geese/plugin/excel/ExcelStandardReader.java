package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.mapping.ClientMapping;
import com.geese.plugin.excel.util.Assert;

import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

/**
 * Created by Administrator on 2017/3/12.
 *
 */
public class ExcelStandardReader {
    // 处理客户端输入信息
    private ClientMapping clientMapping = new ClientMapping();

    public static ExcelStandardReader newInstance(InputStream excelInput) {
        ExcelStandardReader instance = new ExcelStandardReader();
        instance.getClientMapping().setExcelInput(excelInput);
        return instance;
    }

    public ExcelStandardReader query(String query) {
        Assert.notEmpty(query);
        clientMapping.getQueries().add(query);
        return this;
    }

    public ExcelStandardReader query(String first, String second, String... more) {
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

    public ClientMapping getClientMapping() {
        return clientMapping;
    }

    public void setClientMapping(ClientMapping clientMapping) {
        this.clientMapping = clientMapping;
    }
}
