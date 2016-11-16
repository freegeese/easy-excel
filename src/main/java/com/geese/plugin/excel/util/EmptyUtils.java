package com.geese.plugin.excel.util;

import java.lang.reflect.Array;
import java.util.Map;

/**
 * 空值判断工具
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:30
 * @sine 0.0.1
 */
public class EmptyUtils {

    /**
     * 断言为Null
     *
     * @param value
     * @return
     */
    public static boolean isNull(Object value) {
        return null == value;
    }

    /**
     * 断言不为Null
     *
     * @param value
     * @return
     */
    public static boolean notNull(Object value) {
        return null != value;
    }

    /**
     * 断言为空
     *
     * @param value
     * @return
     */
    public static boolean isEmpty(Object value) {
        return !notEmpty(value);
    }

    /**
     * 断言不为空
     *
     * @param value
     * @return
     */
    public static boolean notEmpty(Object value) {
        if (isNull(value)) {
            return false;
        }
        // 值类型
        Class<?> valueClass = value.getClass();
        // 为字符串类型
        if (CharSequence.class.isAssignableFrom(valueClass)) {
            return value.toString().trim().length() > 0;
        }
        // 集合类型
        if (Iterable.class.isAssignableFrom(valueClass)) {
            return ((Iterable) value).iterator().hasNext();
        }
        // Map类型
        if (Map.class.isAssignableFrom(valueClass)) {
            return !((Map) value).isEmpty();
        }
        // 数组类型
        if (valueClass.isArray()) {
            return Array.getLength(value) > 0;
        }
        return true;
    }
}
