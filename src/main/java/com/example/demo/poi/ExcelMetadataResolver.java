package com.example.demo.poi;

import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import org.springframework.stereotype.Component;

@Component
public class ExcelMetadataResolver {

    public List<ExcelColumnMeta> resolve(Class<?> clazz) {

        return Arrays.stream(clazz.getDeclaredFields())
            .filter(f -> f.isAnnotationPresent(ExcelColumn.class))
            .map(field -> {
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                field.setAccessible(true);

                return new ExcelColumnMeta(
                    field,
                    column.header(),
                    column.order(),
                    column.width()
                );
            })
            .sorted(Comparator.comparingInt(ExcelColumnMeta::order))
            .toList();
    }

    public String resolveSheetName(Class<?> clazz) {
        ExcelSheet sheet = clazz.getAnnotation(ExcelSheet.class);
        return sheet != null ? sheet.name() : "Sheet1";
    }
}

