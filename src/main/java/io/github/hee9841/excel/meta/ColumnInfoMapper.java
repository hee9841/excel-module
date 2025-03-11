package io.github.hee9841.excel.meta;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.exception.FailedExcelColumnMappingException;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Workbook;

public class ColumnInfoMapper {

    private final Workbook wb;
    private final Class<?> type;


    private final Map<Integer, ColumnInfo> columnInfoMap;

    private ColumnIndexStrategy columnIndexStrategy;

    private ColumnInfoMapper(Class<?> type, Workbook wb) {
        this.wb = wb;
        this.type = type;
        columnInfoMap = new HashMap<>();
    }

    public static ColumnInfoMapper of(Class<?> type, Workbook workbook) {
        return new ColumnInfoMapper(type, workbook);
    }

    public Map<Integer, ColumnInfo> map() {
        parseExcelAnnotation();
        int autoColumnIndexCnt = 0;

        for (Field field : FieldUtils.getAllFields(type)) {
            if (!field.isAnnotationPresent(ExcelColumn.class)) {
                continue;
            }

            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            field.setAccessible(true);

            int columnIndex = columnIndexStrategy.isFieldOrder()
                ? autoColumnIndexCnt++
                : excelColumn.columnIndex();
            validateColumnIndex(columnIndex, field.getName());

            columnInfoMap.put(columnIndex,
                ColumnInfo.of(field.getName(), excelColumn.headerName())
            );
        }

        return columnInfoMap;
    }

    private void parseExcelAnnotation() {
        if (!type.isAnnotationPresent(Excel.class)) {
            throw new FailedExcelColumnMappingException(
                "Missing the @Excel annotation.", type.getName()
            );
        }
        Excel excel = type.getAnnotation(Excel.class);
        columnIndexStrategy = excel.columnIndexStrategy();
    }

    private void validateColumnIndex(int columnIndex, String fieldName) {
        //1. Check columnIndex value is negative.
        if (columnIndex < 0) {
            throw new FailedExcelColumnMappingException(String.format(
                "Invalid column index : The column index of field '%s' is negative or "
                    + "column index value was not specified when column index strategy is USER_DEFINED.\n"
                    + "Please Change index value to non-negative or Use 'FIELD_ORDER' strategy."
                , fieldName),
                type.getName()
            );
        }

        // 2. Check the columnIndex contains in columnInfoMap
        if (columnInfoMap.containsKey(columnIndex)) {
            throw new FailedExcelColumnMappingException(String.format(
                "Invalid column index : Duplicate value(%d) detected in fields (%s, %s). "
                , columnIndex, fieldName, columnInfoMap.get(columnIndex).getFieldName()),
                type.getName()
            );
        }
    }

}
