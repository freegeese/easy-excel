package com.geese.plugin.excel;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;

/**
 * Created by Administrator on 2017/3/11.
 */
public class ExcelOperationsProxy implements InvocationHandler {
    // 目标对象
    private ExcelOperations target;

    // 构建代理对象的时候织入目标对象
    public ExcelOperationsProxy(ExcelOperations target) {
        this.target = target;
    }

    @Override
    public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
        System.out.println("调用之前：" + method);
        Object returnValue = method.invoke(target, args);
        System.out.println("调用之后：" + returnValue);
        return returnValue;
    }
}
