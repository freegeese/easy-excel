package com.geese.plugin.excel.config;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collection;

/**
 * Excel 配置信息
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 15:53
 * @sine 0.0.1
 */
public class ExcelConfig {
    /**
     * 读取Excel所需的输入源
     */
    private InputStream input;

    /**
     * 生成Excel后输出到哪里
     */
    private OutputStream output;

    /**
     * 生成Excel需要用到的模板
     */
    private InputStream template;

    /**
     * Excel中Sheet表格的配置信息
     */
    private Collection<SheetConfig> sheetConfigs;

    public InputStream getInput() {
        return input;
    }

    public void setInput(InputStream input) {
        this.input = input;
    }

    public OutputStream getOutput() {
        return output;
    }

    public void setOutput(OutputStream output) {
        this.output = output;
    }

    public InputStream getTemplate() {
        return template;
    }

    public void setTemplate(InputStream template) {
        this.template = template;
    }

    public Collection<SheetConfig> getSheetConfigs() {
        return sheetConfigs;
    }

    public void setSheetConfigs(Collection<SheetConfig> sheetConfigs) {
        this.sheetConfigs = sheetConfigs;
    }
}
