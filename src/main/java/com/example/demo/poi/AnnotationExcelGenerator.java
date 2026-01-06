package com.example.demo.poi;

import java.lang.reflect.Field;
import java.util.List;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;

@Component
@RequiredArgsConstructor
public class AnnotationExcelGenerator {

    private final ExcelMetadataResolver resolver;

    public <T> Workbook generate(List<T> data, Class<T> clazz) {

        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        String sheetName = resolver.resolveSheetName(clazz);
        Sheet sheet = workbook.createSheet(sheetName);

        List<ExcelColumnMeta> columns = resolver.resolve(clazz);

        createHeader(sheet, columns);
        createBody(sheet, columns, data);

        return workbook;
    }

    private void createHeader(Sheet sheet, List<ExcelColumnMeta> columns) {
        Row header = sheet.createRow(0);

        for (int i = 0; i < columns.size(); i++) {
            ExcelColumnMeta meta = columns.get(i);
            Cell cell = header.createCell(i);
            cell.setCellValue(meta.header());
            sheet.setColumnWidth(i, meta.width());
        }
    }

    private <T> void createBody(
        Sheet sheet,
        List<ExcelColumnMeta> columns,
        List<T> data
    ) {
        int rowIdx = 1;

        for (T item : data) {
            Row row = sheet.createRow(rowIdx++);

            for (int col = 0; col < columns.size(); col++) {
                ExcelColumnMeta meta = columns.get(col);
                Object value = getValue(meta.field(), item);
                row.createCell(col).setCellValue(valueToString(value));
            }
        }
    }

    private Object getValue(Field field, Object target) {
        try {
            return field.get(target);
        } catch (IllegalAccessException e) {
            throw new IllegalStateException(e);
        }
    }

    private String valueToString(Object value) {
        return value == null ? "" : value.toString();
    }
}

