package com.geese.plugin.excel.util;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Bean属性操作工具
 * <p>主要提供：bean之间的属性复制，map to bean的转换，bean to map的转换<p/>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:28
 * @sine 0.0.1
 */
public class BeanPropertyUtils {

    private static SimpleDateFormat sdf = new SimpleDateFormat();

    /**
     * null -> ""
     *
     * @param bean
     */
    public static void nullToEmpty(Object bean) {
        List<Field> fields = getDeclaredFields(bean.getClass(), true);
        for (Field field : fields) {
            try {
                field.setAccessible(true);
                Class fieldType = field.getType();
                if (String.class == fieldType) {
                    Object val = field.get(bean);
                    if (null == val) {
                        field.set(bean, "");
                    }
                }
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * Map转换为Bean
     *
     * @param map
     * @param bean
     * @return
     */
    public static Object mapToBean(Map map, Object bean) {
        List<Field> fields = getDeclaredFields(bean.getClass(), true);
        Map<String, Field> fieldNameMap = new HashMap<>();
        for (Field field : fields) {
            fieldNameMap.put(field.getName(), field);
        }
        Set keys = map.keySet();
        for (Object key : keys) {
            if (fieldNameMap.containsKey(key)) {
                Field field = fieldNameMap.get(key);
                field.setAccessible(true);
                try {
                    setFieldValue(bean, field, map.get(key));
                } catch (IllegalAccessException | NoSuchFieldException e) {
                    e.printStackTrace();
                } finally {
                    field.setAccessible(false);
                }
            }
        }
        return bean;
    }

    /**
     * 设置字段的值
     *
     * @param target 目标对象
     * @param field  字段
     * @param value  字段值
     * @return 目标对象
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     */
    private static Object setFieldValue(Object target, Field field, Object value) throws NoSuchFieldException, IllegalAccessException {
        notNull(target);
        notNull(field);

        if (value == null) {
            return target;
        }
        // 字段类型
        field.setAccessible(true);
        Class fieldType = field.getType();

        // 值不是String类型，直接设置
        if (String.class != value.getClass()) {
            field.set(target, value);
            return target;
        }

        // 值是String类型，可能存在多种情况
        String valueString = value.toString().trim();

        // 字段为String类型
        if (String.class == fieldType) {
            field.set(target, valueString);
            return target;
        }

        // 字段为基本类型
        if (int.class == fieldType || Integer.class == fieldType) {
            field.set(target, Integer.valueOf(valueString));
            return target;
        }
        if (long.class == fieldType || Long.class == fieldType) {
            field.set(target, Long.valueOf(valueString));
            return target;
        }
        if (double.class == fieldType || Double.class == fieldType) {
            field.set(target, Double.valueOf(valueString));
            return target;
        }
        if (float.class == fieldType || Float.class == fieldType) {
            field.set(target, Float.valueOf(valueString));
            return target;
        }
        if (short.class == fieldType || Short.class == fieldType) {
            field.set(target, Short.valueOf(valueString));
            return target;
        }
        if (byte.class == fieldType || Byte.class == fieldType) {
            field.set(target, Short.valueOf(valueString));
            return target;
        }
        if (Boolean.class == fieldType || boolean.class == fieldType) {
            field.set(target, Boolean.valueOf(valueString));
            return target;
        }

        // 字段为大数字类型
        if (BigDecimal.class == fieldType) {
            BigDecimal.valueOf(Double.valueOf(valueString));
            return target;
        }
        if (BigInteger.class == fieldType) {
            BigInteger.valueOf(Long.valueOf(valueString));
            return target;
        }

        // 字段为日期类型
        if (Date.class == fieldType) {
            String dateString = valueString.replaceAll("[^\\d]+", "");

            switch (dateString.length()) {
                case 4:
                    sdf.applyPattern("yyyy");
                    break;
                case 6:
                    sdf.applyPattern("yyyyMM");
                    break;
                case 8:
                    sdf.applyPattern("yyyyMMdd");
                    break;
                case 10:
                    sdf.applyPattern("yyyyMMddHH");
                    break;
                case 12:
                    sdf.applyPattern("yyyyMMddHHmm");
                    break;
                case 14:
                    sdf.applyPattern("yyyyMMddHHmmss");
                    break;
                case 17:
                    sdf.applyPattern("yyyyMMddHHmmssSSS");
                    break;
                default:
                    throw new IllegalArgumentException("不支持的日期格式：" + valueString);
            }
            try {
                field.set(target, sdf.parse(dateString));
            } catch (ParseException e) {
                e.printStackTrace();
            }
            return target;
        }

        // 字段为枚举类型
        if (fieldType.isEnum()) {
            field.set(target, Enum.valueOf(fieldType, valueString));
            return target;
        }

        throw new IllegalArgumentException("不支持的字段类型处理：" + fieldType);
    }


    /**
     * 复制属性到一个Map
     *
     * @param bean 来源对象
     * @return 目标Map
     */
    public static Map beanToMap(Object bean) {
        return beanToMap(bean, (List<String>) null);
    }

    /**
     * 复制属性到一个Map
     *
     * @param bean     来源对象
     * @param includes 需要复制的属性
     * @return 目标Map
     */
    public static Map beanToMap(Object bean, List<String> includes) {
        return (Map) copy(bean, new HashMap(), includes, null);
    }

    /**
     * 复制属性到一个Map
     *
     * @param bean     来源对象
     * @param excludes 不需要复制的属性
     * @return 目标Map
     */
    public static Map beanToMap(Object bean, String... excludes) {
        return (Map) copy(bean, new HashMap(), null, Arrays.asList(excludes));
    }

    /**
     * 复制属性到一个Map
     *
     * @param beans 来源对象
     * @return 目标Map
     */
    public static List<Map> beanToMap(List beans) {
        return beanToMap(beans, (List<String>) null);
    }

    /**
     * 复制属性到一个Map
     *
     * @param beans    来源对象
     * @param includes 需要复制的属性
     * @return 目标Map集合
     */
    public static List<Map> beanToMap(List beans, List<String> includes) {
        List<Map> toList = new ArrayList<>();
        for (Object o : beans) {
            toList.add(beanToMap(o, includes));
        }
        return toList;
    }

    /**
     * 复制属性到一个Map
     *
     * @param beans    来源对象
     * @param excludes 需要排除的属性
     * @return 目标Map集合
     */
    public static List<Map> beanToMap(List beans, String... excludes) {
        List<Map> toList = new ArrayList<>();
        for (Object o : beans) {
            toList.add(beanToMap(o, excludes));
        }
        return toList;
    }

    /**
     * 对象属性复制
     *
     * @param from 来源对象
     * @param to   目标对象
     */
    public static Object copy(Object from, Object to) {
        return copy(from, to, null, null);
    }

    /**
     * 对象属性复制
     *
     * @param from     来源对象
     * @param to       目标对象
     * @param includes 指定需要复制的属性
     */
    public static Object copy(Object from, Object to, List<String> includes) {
        return copy(from, to, includes, null);
    }

    /**
     * 对象属性复制
     *
     * @param from     来源对象
     * @param to       目标对象
     * @param excludes 排除不需要复制的属性
     */
    public static Object copy(Object from, Object to, String... excludes) {
        return copy(from, to, null, Arrays.asList(excludes));
    }

    /**
     * 对象属性复制
     *
     * @param from 来源对象
     * @param to   目标类
     * @param <T>  目标对象类型
     * @return 目标对象
     */
    public static <T> T copy(Object from, Class<T> to) {
        return copy(from, to, (List<String>) null);
    }

    /**
     * 对象属性复制
     *
     * @param from     来源对象
     * @param to       目标类
     * @param excludes 指定不需要复制的属性
     * @param <T>      目标对象类型
     * @return 目标对象
     */
    public static <T> T copy(Object from, Class<T> to, String... excludes) {
        T t = null;
        try {
            t = to.newInstance();
            copy(from, t, null, Arrays.asList(excludes));
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        return t;
    }

    /**
     * 对象属性复制
     *
     * @param from     来源对象
     * @param to       目标类
     * @param includes 指定需要复制的属性
     * @param <T>      目标对象类型
     * @return 目标对象
     */
    public static <T> T copy(Object from, Class<T> to, List<String> includes) {
        T t = null;
        try {
            t = to.newInstance();
            copy(from, t, includes, null);
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        return t;
    }

    /**
     * 对象属性复制
     *
     * @param from 来源对象集合
     * @param to   目标类
     * @param <T>  目标对象类型
     * @return 目标对象集合
     */
    public static <T> List<T> copy(List<?> from, Class<T> to) {
        return copy(from, to, (List<String>) null);
    }

    /**
     * 对象属性复制
     *
     * @param from     来源对象集合
     * @param to       目标类
     * @param excludes 制定不需要复制的属性
     * @param <T>      目标对象类型
     * @return 目标对象集合
     */
    public static <T> List<T> copy(List from, Class<T> to, String... excludes) {
        List<T> toList = new ArrayList<>();
        for (Object o : from) {
            toList.add(copy(o, to, excludes));
        }
        return toList;
    }

    /**
     * 对象属性复制
     *
     * @param from     来源对象
     * @param to       目标对象
     * @param includes 指定需要复制的属性
     * @param <T>      目标对象类型
     * @return 目标对象集合
     */
    public static <T> List<T> copy(List<?> from, Class<T> to, List<String> includes) {
        List<T> toList = new ArrayList<>();
        for (Object o : from) {
            toList.add(copy(o, to, includes));
        }
        return toList;
    }

    /**
     * 属性复制
     *
     * @param from     来源对象
     * @param to       目标对象
     * @param includes 指定需要复制的属性（includes 优先于 excludes）
     * @param excludes 指定不需要复制的属性
     */
    private static Object copy(Object from, Object to, List<String> includes, List<String> excludes) {
        notNull(from);
        notNull(to);

        // 检查是否有需要导入或排除的属性
        boolean hasIncludes = (null != includes && !includes.isEmpty());
        boolean hasExcludes = false;
        if (!hasIncludes) {
            hasExcludes = (null != excludes && !excludes.isEmpty());
        }

        // 来源字段
        Class<?> fromClass = from.getClass();
        List<Field> fromFields = getDeclaredFields(fromClass, true);

        // 目标字段
        Class<?> toClass = to.getClass();
        boolean isMap = Map.class.isAssignableFrom(toClass);
        List<String> toFieldNames = null;
        if (!isMap) {
            toFieldNames = new ArrayList<>();
            List<Field> toFields = getDeclaredFields(toClass, true);
            for (Field toField : toFields) {
                toFieldNames.add(toField.getName());
            }
        }

        // 来源字段名称
        String fromFieldName;
        for (Field fromField : fromFields) {
            fromFieldName = fromField.getName();
            try {
                fromField.setAccessible(true);
                // 目标对象是Map类型
                if (isMap) {
                    ((Map) to).put(fromFieldName, fromField.get(from));
                    continue;
                }
                // 目标对象是Bean类型
                if (toFieldNames.contains(fromFieldName)) {
                    // 有需要导入的
                    if (hasIncludes && !includes.contains(fromFieldName)) {
                        continue;
                    }
                    // 有需要排除的
                    if (hasExcludes && excludes.contains(fromFieldName)) {
                        continue;
                    }
                    fromField.set(to, fromField.get(from));
                }
            } catch (IllegalAccessException e) {
                System.err.println("属性访问异常：" + e.getMessage());
                // e.printStackTrace();
            } finally {
                fromField.setAccessible(false);
            }
        }

        return to;
    }

    /**
     * 获取已经声明的字段
     *
     * @param aClass 目标类型
     * @param deep   是否深度获取(向上查找父类的字段)
     * @return 获取的字段集合
     */
    private static List<Field> getDeclaredFields(Class aClass, boolean deep) {
        List<Field> fieldsStore = new ArrayList<>();
        getFields(fieldsStore, aClass, true, deep);
        return fieldsStore;
    }

    /**
     * 获取字段
     *
     * @param findedFields 存放获取到字段的容器
     * @param aClass       目标类
     * @param declared     字段是否声明
     * @param deep         是否深度查找
     */
    private static void getFields(List<Field> findedFields, Class aClass, boolean declared, boolean deep) {
        if (Object.class != aClass) {
            Field[] fields = declared ? aClass.getDeclaredFields() : aClass.getFields();
            findedFields.addAll(Arrays.asList(fields));
            if (deep) {
                getFields(findedFields, aClass.getSuperclass(), declared, deep);
            }
        }
    }


    private static void notNull(Object object) {
        notNull(object, "[Assertion failed] - this argument is required; it must not be null");
    }

    private static void notNull(Object object, String message) {
        if (object == null) {
            throw new IllegalArgumentException(message);
        }
    }

    public static void main(String[] args) {
        String s = "sd23f3f2d423".replaceAll("[^\\d]+", "");
        System.out.println(s);
    }
}
