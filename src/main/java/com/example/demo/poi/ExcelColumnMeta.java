package com.example.demo.poi;

import java.lang.reflect.Field;

public record ExcelColumnMeta(
    Field field,
    String header,
    int order,
    int width
) {

    public static ExcelColumnMeta from(Field field, String header, int order, int width) {
        return new ExcelColumnMeta(field, header, order, width);
    }
}
