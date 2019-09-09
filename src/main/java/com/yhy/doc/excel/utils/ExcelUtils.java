package com.yhy.doc.excel.utils;

import lombok.extern.slf4j.Slf4j;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.*;

/**
 * author : 颜洪毅
 * e-mail : yhyzgn@gmail.com
 * time   : 2019-09-09 12:14
 * version: 1.0.0
 * desc   :
 */
@Slf4j
public class ExcelUtils {

    private ExcelUtils() {
        throw new UnsupportedOperationException("Utils class can not be instantiate.");
    }

//    public static void read(File file) {
//        Workbook workbook = WorkbookFactory.create(file);
//    }
//
//    public static void read(MultipartFile file) {
//        Workbook workbook = WorkbookFactory.create(file);
//    }
//
//    public static void read(InputStream is) {
//        Workbook workbook = WorkbookFactory.create(is);
//    }

    public static <T> T instantiate(Class<T> clazz) {
        try {
            return instantiate(clazz.getConstructor());
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static <T> T instantiate(Constructor<T> constructor) {
        try {
            return constructor.newInstance();
        } catch (InstantiationException | IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static Method setter(Field field, Class<?> clazz) throws IntrospectionException {
        PropertyDescriptor descriptor = new PropertyDescriptor(field.getName(), clazz);
        return descriptor.getWriteMethod();
    }

    public static Method getter(Field field, Class<?> clazz) throws IntrospectionException {
        PropertyDescriptor descriptor = new PropertyDescriptor(field.getName(), clazz);
        return descriptor.getReadMethod();
    }

    public static Type getParamType(Class<?> clazz, Class<?> interfaceClazz, int index) {
        Type[] types = clazz.getGenericInterfaces();
        ParameterizedType ptp;
        Type raw;
        for (Type tp : types) {
            if (tp instanceof ParameterizedType) {
                ptp = (ParameterizedType) tp;
                raw = ptp.getRawType();
                if (raw.getTypeName().equalsIgnoreCase(interfaceClazz.getName()) && ptp.getActualTypeArguments().length > index) {
                    // 匹配到对应的接口
                    return ptp.getActualTypeArguments()[index];
                }
            }
        }
        return null;
    }
}
