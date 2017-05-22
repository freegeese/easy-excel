package com.geese.plugin.excel;

import com.geese.plugin.excel.filter.Filter;
import com.geese.plugin.excel.filter.WriteFilter;
import com.geese.plugin.excel.mapping.ClientMapping;
import com.geese.plugin.excel.mapping.ExcelMapping;
import com.geese.plugin.excel.util.Assert;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * 读取Excel
 */
public class ExcelWriter {
    // 处理客户端输入信息
    private ClientMapping clientMapping = new ClientMapping();

    public static ExcelWriter newInstance(OutputStream excelOutput) {
        ExcelWriter instance = new ExcelWriter();
        instance.clientMapping.setExcelOutput(excelOutput);
        return instance;
    }

    public ExcelWriter useXlsxFormat() {
        clientMapping.setUseXlsFormat(false);
        return this;
    }

    public ExcelWriter setTemplate(File template) {
        clientMapping.setExcelOutputTemplate(template);
        return this;
    }

    public ExcelWriter insert(String insert) {
        Assert.notEmpty(insert);
        clientMapping.getInserts().add(insert);
        return this;
    }

    public ExcelWriter insert(String first, String second, String... more) {
        Assert.notEmpty(first, second);
        List<String> inserts = clientMapping.getInserts();
        inserts.add(first);
        inserts.add(second);
        if (null != more && more.length > 0) {
            inserts.addAll(Arrays.asList(more));
        }
        return this;
    }

    public ExcelWriter addData(List<Map> tableData, String switchSheet) {
        clientMapping.addTableData(tableData, switchSheet);
        return this;
    }

    public ExcelWriter addData(Map pointData, String switchSheet) {
        clientMapping.addPointData(pointData, switchSheet);
        return this;
    }

    public ExcelWriter addValidation(ExcelValidation validation, String switchSheet) {
        clientMapping.addValidation(validation, switchSheet);
        return this;
    }

    public ExcelWriter addValidations(Collection<ExcelValidation> validations, String switchSheet) {
        clientMapping.addValidations(validations, switchSheet);
        return this;
    }


    public ExcelWriter filter(WriteFilter filter, String switchSheet) {
        clientMapping.addFilter(filter, switchSheet);
        return this;
    }

    public ExcelWriter filters(WriteFilter[] filters, String switchSheet) {
        return filters(Arrays.asList(filters), switchSheet);
    }

    public ExcelWriter filters(Collection<WriteFilter> filters, String switchSheet) {
        for (WriteFilter filter : filters) {
            clientMapping.addFilter(filter, switchSheet);
        }
        return this;
    }

    public ExcelResult execute() throws IOException, InvalidFormatException {
        // 创建Workbook
        File template = clientMapping.getExcelOutputTemplate();
        Workbook workbook = null;
        if (null != template) {
            workbook = WorkbookFactory.create(template);
        }
        if (null == workbook) {
            workbook = clientMapping.getUseXlsFormat() ? new HSSFWorkbook() : new XSSFWorkbook();
        }
        // Excel 数据校验
        Map<String, Collection<ExcelValidation>> sheetAndValidationMap = clientMapping.getSheetAndValidationMap();
        if (!sheetAndValidationMap.isEmpty()) {
            for (Map.Entry<String, Collection<ExcelValidation>> entry : sheetAndValidationMap.entrySet()) {
                String key = entry.getKey();
                Sheet sheet = workbook.getSheet(key);
                if (null == sheet && ExcelHelper.isNumber(key)) {
                    sheet = workbook.getSheetAt(Integer.valueOf(key));
                }
                Collection<ExcelValidation> validations = entry.getValue();
                for (ExcelValidation validation : validations) {
                    validation.addToSheet(sheet);
                }
            }
        }
        // 如果两者都为空，则不进行读写操作
        if (clientMapping.getInserts().isEmpty() && clientMapping.getQueries().isEmpty()) {
            workbook.write(clientMapping.getExcelOutput());
            return null;
        }

        // 把客户输入转换为Excel映射信息
        ExcelMapping excelMapping = clientMapping.parseClientInput();

        // Excel操作接口代理
        ExcelOperations proxy = ExcelOperationsProxyFactory.getProxy();
        proxy.writeExcel(workbook, excelMapping);
        workbook.write(clientMapping.getExcelOutput());
        ExcelResult excelResult = new ExcelResult();
        excelResult.setContext(ExcelTemplate.getContext());
        return excelResult;
    }
}
