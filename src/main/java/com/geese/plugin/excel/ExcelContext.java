package com.geese.plugin.excel;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Created by Administrator on 2017/6/29.
 */
public abstract class ExcelContext {

    private static final ThreadLocal<Map> LOCAL_CONTEXT = new ThreadLocal<Map>();

    public static Map get() {
        if (null == LOCAL_CONTEXT.get()) {
            LOCAL_CONTEXT.set(new LinkedHashMap<>());
        }
        return LOCAL_CONTEXT.get();
    }

    public static void set(Map context) {
        LOCAL_CONTEXT.set(context);
    }

    public static void remove() {
        LOCAL_CONTEXT.remove();
    }

}
