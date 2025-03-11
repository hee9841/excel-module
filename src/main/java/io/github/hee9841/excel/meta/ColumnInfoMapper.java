package io.github.hee9841.excel.meta;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumStyle;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.exception.ExcelStyleException;
import io.github.hee9841.excel.strategy.CellTypeStrategy;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import io.github.hee9841.excel.style.ExcelCellStyle;
import io.github.hee9841.excel.style.NoCellStyle;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.Map;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class ColumnInfoMapper {

    private static final String STANDARD_STYLE = NoCellStyle.class.getName();

    private final Workbook wb;
    private final Class<?> type;

    private final CellStyle defaultHeaderStyle;
    private final CellStyle defaultBodyStyle;

    private ColumnIndexStrategy columnIndexStrategy;
    private CellTypeStrategy cellTypeStrategy;

    private ColumnInfoMapper(Class<?> type, Workbook wb) {
        this.wb = wb;
        this.type = type;
        this.defaultHeaderStyle = wb.createCellStyle();
        this.defaultBodyStyle = wb.createCellStyle();
    }

    public static ColumnInfoMapper of(Class<?> type, Workbook workbook) {
        return new ColumnInfoMapper(type, workbook);
    }

    public Map<Integer, ColumnInfo> map() {
        parsingExcelAnnotation();
        return parsingExcelColumns();
    }

    private void parsingExcelAnnotation() {
        if (!type.isAnnotationPresent(Excel.class)) {
            throw new ExcelException(
                "Missing the @Excel annotation.", type.getName()
            );
        }
        Excel excel = type.getAnnotation(Excel.class);
        columnIndexStrategy = excel.columnIndexStrategy();
        cellTypeStrategy = excel.cellTypeStrategy();

        //set default style
        getExcelCellStyle(excel.defaultHeaderStyle()).apply(defaultHeaderStyle);
        getExcelCellStyle(excel.defaultBodyStyle()).apply(defaultBodyStyle);
    }

    private Map<Integer, ColumnInfo> parsingExcelColumns() {
        int autoColumnIndexCnt = 0;
        Map<Integer, ColumnInfo> result = new HashMap<>();

        for (Field field : FieldUtils.getAllFields(type)) {
            if (!field.isAnnotationPresent(ExcelColumn.class)) {
                continue;
            }

            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            field.setAccessible(true);

            //set column index value
            int columnIndex = columnIndexStrategy.isFieldOrder()
                ? autoColumnIndexCnt++
                : excelColumn.columnIndex();
            validateColumnIndex(result, columnIndex, field.getName());

            //get column info
            result.put(columnIndex,
                getColumnInfo(excelColumn, field.getType(), field.getName()));
        }

        return result;
    }

    private void validateColumnIndex(Map<Integer, ColumnInfo> columnInfoMap, int columnIndex,
        String fieldName) {
        //1. Check columnIndex value is negative.
        if (columnIndex < 0) {
            throw new ExcelException(String.format(
                "Invalid column index : The column index of field '%s' is negative or "
                    + "column index value was not specified when column index strategy is USER_DEFINED.\n"
                    + "Please Change index value to non-negative or Use 'FIELD_ORDER' strategy."
                , fieldName),
                type.getName()
            );
        }

        // 2. Check the columnIndex contains in columnInfoMap
        if (columnInfoMap.containsKey(columnIndex)) {
            throw new ExcelException(String.format(
                "Invalid column index : Duplicate value(%d) detected in fields (%s, %s). "
                , columnIndex, fieldName, columnInfoMap.get(columnIndex).getFieldName()),
                type.getName()
            );
        }
    }

    private ColumnInfo getColumnInfo(ExcelColumn excelColumn, Class<?> fieldType,
        String fieldName) {
        //Get cell type
        CellType cellType = getCellType(excelColumn.columnCellType(), fieldType, fieldName);

        //Set Cell style
        CellStyle headerStyle = updateCellStyle(excelColumn.headerStyle(), defaultHeaderStyle);
        CellStyle bodyStyle = updateCellStyle(excelColumn.bodyStyle(), defaultBodyStyle);

        return ColumnInfo.of(
            fieldName,
            excelColumn.headerName(),
            cellType,
            headerStyle,
            bodyStyle
        );
    }

    private CellType getCellType(CellType columnCellType, Class<?> fieldType, String fieldName) {
        // When cell type strategy is AUTO and column cell type is not specified
        // or when column cell type is AUTO,
        // the column's cell type is automatically determined based on the field type.
        if ((cellTypeStrategy.isAuto() && columnCellType.isNone()) || columnCellType.isAuto()) {
            return CellType.from(fieldType);
        }

        // When the specified column cell type is not compatible with the field type, throw an exception.
        if (!columnCellType.equals(CellType.findMatchingCellType(fieldType, columnCellType))) {
            throw new ExcelException(String.format(
                "Invalid cell type : The cell type of field '%s' is not compatible with the specified cell type(%s).",
                fieldName,
                columnCellType.name()
            ), type.getName());
        }

        return columnCellType;
    }

    private CellStyle updateCellStyle(ExcelColumStyle style, CellStyle defaultStyle) {

        CellStyle cellStyle = wb.createCellStyle();
        //todo docs NoStyle(default) 같은 경우 우선 순위 나중, NoStyle로 Column에 우선순위로 적요하고 싶으면 default 값이 아닌,
        // 사용자 정의해서 적용해야함

        // When cell style is default value of @ExcelColumn,
        // return the default cell style specified by @Excel.
        if (style.cellStyleClass().getName().equals(STANDARD_STYLE)) {
            cellStyle.cloneStyleFrom(defaultStyle);
            return cellStyle;
        }

        // When cell style is not default value of @ExcelColumn,
        // apply and return the cell style specified by @ExcelColumn.
        getExcelCellStyle(style).apply(cellStyle);
        return cellStyle;
    }

    @SuppressWarnings("unchecked")
    private <E> ExcelCellStyle getExcelCellStyle(ExcelColumStyle excelColumStyle) {
        Class<? extends ExcelCellStyle> cellStyleClass = excelColumStyle.cellStyleClass();
        //1. case of enum
        if (cellStyleClass.isEnum()) {
            try {
                Class<? extends Enum> enumClass = cellStyleClass.asSubclass(Enum.class);
                return (ExcelCellStyle) Enum.valueOf(enumClass, excelColumStyle.enumName());
            } catch (IllegalArgumentException e) {
                throw new ExcelStyleException(
                    String.format(
                        "Failed to convert Enum cell style to cellStyle : "
                            + "Enum value '%s' not found in style class '%s'.",
                        excelColumStyle.enumName(), cellStyleClass.getName()), type.getName(), e);
            }
        }

        //2. case of class
        try {
            return cellStyleClass.getDeclaredConstructor().newInstance();
        } catch (NoSuchMethodException | IllegalAccessException |
                 InstantiationException | InvocationTargetException e
        ) {
            throw new ExcelStyleException(
                String.format("Failed to instantiate cellStyle class of %s.",
                    cellStyleClass.getName()), type.getName(), e);
        }
    }

}
