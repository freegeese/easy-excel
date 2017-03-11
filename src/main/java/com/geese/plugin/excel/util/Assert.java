package com.geese.plugin.excel.util;

import java.util.Arrays;
import java.util.List;

/**
 * 空值检查
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:29
 * @sine 0.0.1
 */
public final class Assert {
    private Assert() {
    }

    public static <T> T notNull(T reference) {
        if (reference == null) {
            throw new NullPointerException();
        }
        return reference;
    }

    public static <T> T notNull(T reference, Object errorMessage) {
        if (reference == null) {
            throw new NullPointerException(String.valueOf(errorMessage));
        }
        return reference;
    }

    public static <T> T notNull(
            T reference, String errorMessageTemplate, Object... errorMessageArgs) {
        if (reference == null) {
            throw new NullPointerException(format(errorMessageTemplate, errorMessageArgs));
        }
        return reference;
    }

    public static void notNull(Object first, Object second, Object... more) {
        List<Object> args = Arrays.asList(first, second, more);
        for (Object arg : args) {
            if (null == arg) {
                throw new NullPointerException();
            }
        }
    }

    public static <T> T notEmpty(T reference) {
        if (EmptyUtils.isEmpty(reference)) {
            throw new IllegalArgumentException();
        }
        return reference;
    }

    public static <T> T notEmpty(T reference, Object errorMessage) {
        if (EmptyUtils.isEmpty(reference)) {
            throw new IllegalArgumentException(String.valueOf(errorMessage));
        }
        return reference;
    }

    public static <T> T notEmpty(
            T reference, String errorMessageTemplate, Object... errorMessageArgs) {
        if (EmptyUtils.isEmpty(reference)) {
            throw new IllegalArgumentException(format(errorMessageTemplate, errorMessageArgs));
        }
        return reference;
    }

    public static void notEmpty(Object first, Object second, Object... more) {
        List<Object> args = Arrays.asList(first, second, more);
        for (Object arg : args) {
            if (EmptyUtils.isEmpty(arg)) {
                throw new IllegalArgumentException();
            }
        }
    }


    static String format(String template, Object... args) {
        template = String.valueOf(template); // null -> "null"
        // start substituting the arguments into the '%s' placeholders
        StringBuilder builder = new StringBuilder(template.length() + 16 * args.length);
        int templateStart = 0;
        int i = 0;
        while (i < args.length) {
            int placeholderStart = template.indexOf("%s", templateStart);
            if (placeholderStart == -1) {
                break;
            }
            builder.append(template, templateStart, placeholderStart);
            builder.append(args[i++]);
            templateStart = placeholderStart + 2;
        }
        builder.append(template, templateStart, template.length());

        if (i < args.length) {
            builder.append(" [");
            builder.append(args[i++]);
            while (i < args.length) {
                builder.append(", ");
                builder.append(args[i++]);
            }
            builder.append(']');
        }

        return builder.toString();
    }
}
