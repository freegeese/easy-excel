package com.geese.plugin.excel;

import java.lang.reflect.Proxy;

/**
 * Excel操作接口代理工厂（单例模式）
 */
public class ExcelOperationsProxyFactory {

    private static ExcelOperations proxy;

    public static ExcelOperations getProxy() {
        if (null != proxy) {
            return proxy;
        }
        synchronized (ExcelOperationsProxyFactory.class) {
            if (null == proxy) {
                synchronized (ExcelOperations.class) {
                    ExcelTemplate excelTemplate = new ExcelTemplate();
                    ExcelOperationsProxy handler = new ExcelOperationsProxy(excelTemplate);
                    Class<? extends ExcelTemplate> targetClass = excelTemplate.getClass();
                    ClassLoader loader = targetClass.getClassLoader();
                    Class<?>[] interfaces = targetClass.getInterfaces();
                    proxy = (ExcelOperations) Proxy.newProxyInstance(loader, interfaces, handler);
                }
            }
            return proxy;
        }
    }

}
