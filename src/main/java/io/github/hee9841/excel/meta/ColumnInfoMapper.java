package io.github.hee9841.excel.meta;

import static io.github.hee9841.excel.global.SystemValues.ALLOWED_FIELD_TYPES;
import static io.github.hee9841.excel.global.SystemValues.ALLOWED_FIELD_TYPES_STRING;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.annotation.ExcelColumnStyle;
import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.exception.ExcelStyleException;
import io.github.hee9841.excel.format.CellFormats;
import io.github.hee9841.excel.format.ExcelDataFormater;
import io.github.hee9841.excel.strategy.CellTypeStrategy;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import io.github.hee9841.excel.strategy.DataFormatStrategy;
import io.github.hee9841.excel.style.ExcelCellStyle;
import io.github.hee9841.excel.style.NoCellStyle;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Modifier;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Maps Java class fields to Excel columns using reflection and annotations.
 * This class processes the {@link io.github.hee9841.excel.annotation.Excel} and
 * {@link io.github.hee9841.excel.annotation.ExcelColumn} annotations to generate column information
 * for Excel export operations. It handles column indexing, cell data types of column determination,
 * and style application based on the annotation configurations.
 *
 * @see ColumnInfo
 * @see ColumnDataType
 * @see Excel
 * @see ExcelColumn
 * @see ExcelColumnStyle
 */
public class ColumnInfoMapper {

    /**
     * The fully qualified class name of the default "no style" class
     */
    private static final String STANDARD_STYLE = NoCellStyle.class.getName();

    /**
     * The Apache POI Workbook to create cell styles for
     */
    private final Workbook wb;
    /**
     * The class type being mapped to Excel
     */
    private final Class<?> type;

    /**
     * The default style to use for header cells
     */
    private final CellStyle defaultHeaderStyle;
    /**
     * The default style to use for body cells
     */
    private final CellStyle defaultBodyStyle;

    /**
     * Strategy for determining column indices
     */
    private ColumnIndexStrategy columnIndexStrategy;
    /**
     * Strategy for determining cell types
     */
    private CellTypeStrategy cellTypeStrategy;
    /**
     * Strategy for determining data formats
     */
    private DataFormatStrategy dataFormatStrategy;


    private ColumnInfoMapper(Class<?> type, Workbook wb) {
        this.wb = wb;
        this.type = type;
        this.defaultHeaderStyle = wb.createCellStyle();
        this.defaultBodyStyle = wb.createCellStyle();
    }

    /**
     * Factory method to create a new {@link ColumnInfoMapper} instance.
     *
     * @param type     The class type to map
     * @param workbook The Apache POI Workbook to create styles for
     * @return A new {@link ColumnInfoMapper} instance
     */
    public static ColumnInfoMapper of(Class<?> type, Workbook workbook) {
        return new ColumnInfoMapper(type, workbook);
    }

    /**
     * Maps the class fields to Excel columns and returns a map of column indices to
     * {@link ColumnInfo} objects.
     * This method processes the {@link io.github.hee9841.excel.annotation.Excel} annotation and
     * all {@link io.github.hee9841.excel.annotation.ExcelColumn} annotations in the class.
     *
     * @return A map of column indices to {@link ColumnInfo} objects
     * @throws ExcelException If the class is not properly annotated or has invalid configuration
     */
    public Map<Integer, ColumnInfo> map() {
        parsingExcelAnnotation();
        return parsingExcelColumns().orElseThrow(() -> new ExcelException(
                String.format("No @ExcelColumn annotations found in class '(%s)'."
                    + " At least one field must be annotated with @ExcelColumn", type.getName())
            )
        );
    }

    /**
     * Parses the {@link Excel} annotation on the class to determine global settings.
     * Sets up the column index strategy, cell type strategy, and data format strategy.
     * Also applies default styles for headers and bodies.
     *
     * @throws ExcelException If the {@link Excel} annotation is missing
     */
    private void parsingExcelAnnotation() {
        validateExcelAnnotation(type);

        Excel excel = type.getAnnotation(Excel.class);
        columnIndexStrategy = excel.columnIndexStrategy();
        cellTypeStrategy = excel.cellTypeStrategy();
        dataFormatStrategy = excel.dataFormatStrategy();

        //set default style
        getExcelCellStyle(excel.defaultHeaderStyle()).apply(defaultHeaderStyle);
        getExcelCellStyle(excel.defaultBodyStyle()).apply(defaultBodyStyle);
    }

    private void validateExcelAnnotation(Class<?> clazz) {
        if (clazz.isInterface()) {
            throw new ExcelException("The class " + clazz.getName()
                + " is interface. You can't annotate interface classes with @Excel");
        }

        if (Modifier.isAbstract(clazz.getModifiers())) {
            throw new ExcelException("The class " + clazz.getName()
                + " is abstract. You can't annotate abstract classes with @Excel");
        }

        if (!type.isAnnotationPresent(Excel.class)) {
            throw new ExcelException(
                "Missing the @Excel annotation.", type.getName()
            );
        }
    }

    /**
     * Parses all fields annotated with {@link io.github.hee9841.excel.annotation.ExcelColumn}
     * in the class and builds a map of
     * column indices to {@link ColumnInfo} objects.
     *
     * @return An Optional containing the map of column indices to {@link ColumnInfo} objects,
     * or an empty Optional
     * if no {@link io.github.hee9841.excel.annotation.ExcelColumn} annotations were found.
     * @throws ExcelException If there are duplicate column indices or other validation errors
     */
    private Optional<Map<Integer, ColumnInfo>> parsingExcelColumns() {
        int autoColumnIndexCnt = 0;
        Map<Integer, ColumnInfo> result = new HashMap<>();

        for (Field field : FieldUtils.getAllFields(type)) {
            if (!field.isAnnotationPresent(ExcelColumn.class)) {
                continue;
            }

            validateField(field);

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

        return result.isEmpty() ? Optional.empty() : Optional.of(result);
    }

    private void validateField(Field field) {

        if (field.getType().isArray()) {
            throw new ExcelException(
                String.format("@ExcelColumn cannot be applied to array type: %s",
                    field.getName()),
                type.getName()
            );
        }

        if (type.isEnum() || type.isPrimitive()) {
            return;
        }

        Class<?> fieldType = field.getType();
        ALLOWED_FIELD_TYPES.stream()
            .filter(allowedType -> allowedType.isAssignableFrom(fieldType))
            .findFirst()
            .orElseThrow(() -> new ExcelException(
                String.format(
                    "%s(%s) Type is %s : @ExcelColumn can only be applied to allowed types(%s).",
                    field.getName(),
                    type.getName(),
                    fieldType.getSimpleName(),
                    ALLOWED_FIELD_TYPES_STRING)
            ));
    }

    /**
     * Validates a column index to ensure it's not negative and not already in use.
     *
     * @param columnInfoMap The current map of column indices to {@link ColumnInfo} objects
     * @param columnIndex   The column index to validate
     * @param fieldName     The name of the field being validated
     * @throws ExcelException If the column index is negative or already in use
     */
    private void validateColumnIndex(Map<Integer, ColumnInfo> columnInfoMap, int columnIndex,
        String fieldName) {
        //1. Check columnIndex value is negative.
        if (columnIndex < 0) {
            throw new ExcelException(String.format(
                "Invalid column index : The column index of '%s' field is negative or "
                    + "column index value was not specified when column index strategy is USER_DEFINED.\n"
                    + "Please Change index value to non-negative or Use 'FIELD_ORDER' strategy."
                , fieldName),
                type.getName()
            );
        }

        // 2. Check the columnIndex contains in columnInfoMap
        if (columnInfoMap.containsKey(columnIndex)) {
            throw new ExcelException(String.format(
                "Invalid column index : Duplicate value(%d) detected in fields (%s, %s)."
                , columnIndex, fieldName, columnInfoMap.get(columnIndex).getFieldName()),
                type.getName()
            );
        }
    }

    /**
     * Creates a {@link ColumnInfo} object for a field based on its {@link ExcelColumn} annotation.
     *
     * @param excelColumn The {@link ExcelColumn} annotation
     * @param fieldType   The type of the field
     * @param fieldName   The name of the field
     * @return A {@link ColumnInfo} object with the appropriate settings
     * @throws ExcelException If the {@link ColumnDataType} is not compatible with the field type
     */
    private ColumnInfo getColumnInfo(ExcelColumn excelColumn, Class<?> fieldType,
        String fieldName) {
        //Get cell type for column
        ColumnDataType columnDataType = getcolumnDataType(excelColumn.columnCellType(), fieldType,
            fieldName);

        //Set Cell style
        CellStyle headerStyle = updateCellStyle(excelColumn.headerStyle(), defaultHeaderStyle);
        CellStyle bodyStyle = updateCellStyle(excelColumn.bodyStyle(), defaultBodyStyle);

        //Set colum cell(body) format
        ExcelDataFormater dataFormater = getDataFormater(excelColumn.format(), columnDataType);
        dataFormater.apply(bodyStyle);

        return ColumnInfo.of(
            fieldName,
            excelColumn.headerName(),
            columnDataType,
            headerStyle,
            bodyStyle
        );
    }

    /**
     * Creates a {@link ExcelDataFormater} for a cell based on the format pattern and
     * {@link ColumnDataType}.
     * Applies automatic formatting if the {@link DataFormatStrategy} is
     * {@link DataFormatStrategy#AUTO_BY_CELL_TYPE} by {@link ColumnDataType}.
     *
     * @param pattern        The format pattern specified in the annotation
     * @param columnDataType The {@link ColumnDataType}
     * @return An {@link ExcelDataFormater} for the cell
     */
    private ExcelDataFormater getDataFormater(String pattern, ColumnDataType columnDataType) {
        // When dataFormatStrategy is "AUTO" and format pattern is "isNone"(empty or null),
        // apply auto format pattern.
        if ((dataFormatStrategy.isAutoByColumnDataType() && CellFormats.isNone(pattern))) {
            pattern = columnDataType.getDataFormatPattern();
        }

        // When dataFormatStrategy is "AUTO" and format pattern is not "isNone"
        // or dataFormatStrategy is "NONE"(format pattern is "isNone" or any value),
        // apply parameter "pattern" value.
        return ExcelDataFormater.of(wb.createDataFormat(), pattern);
    }

    /**
     * Determines the appropriate {@link ColumnDataType} for a field based on the column cell type
     * and the {@link CellTypeStrategy}.
     *
     * @param columnColumnDataType The {@link ColumnDataType} specified in the annotation
     * @param fieldType            The type of the field
     * @param fieldName            The name of the field
     * @return The appropriate {@link ColumnDataType}
     * @throws ExcelException If the specified {@link ColumnDataType} is not compatible with the
     *                        field
     *                        type
     */
    private ColumnDataType getcolumnDataType(ColumnDataType columnColumnDataType,
        Class<?> fieldType, String fieldName) {
        // When cell type strategy is AUTO and column cell type is not specified
        // or when column cell type is AUTO,
        // the column's cell type is automatically determined based on the field type.
        if ((cellTypeStrategy.isAuto() && columnColumnDataType.isNone())
            || columnColumnDataType.isAuto()) {
            return ColumnDataType.from(fieldType);
        }

        // When the specified column cell type is not compatible with the field type, throw an exception.
        if (!columnColumnDataType.equals(
            ColumnDataType.findMatchingCellType(fieldType, columnColumnDataType))) {
            throw new ExcelException(String.format(
                "Invalid cell type : The cell type of '%s' field is not compatible with the specified cell type(%s).",
                fieldName,
                columnColumnDataType.name()
            ), type.getName());
        }

        return columnColumnDataType;
    }

    /**
     * Updates a {@link CellStyle} based on the field's {@link ExcelColumn} annotation.
     * If the annotation specifies the default style, the default style is used.
     * Otherwise, the specified style is applied.
     * <p>
     * Note: {@link io.github.hee9841.excel.style.NoCellStyle} has the lowest priority when applying
     * styles.
     * If you want to give a column the highest priority with no styling, you should create and
     * apply
     * a custom no-style implementation rather than using the default NoCellStyle.(refer to
     * {@link io.github.hee9841.excel.style.NoCellStyle})
     *
     * @param style        The {@link ExcelColumnStyle} defined in the {@link ExcelColumn}
     *                     annotation
     * @param defaultStyle The default {@link CellStyle} to use if no style is specified
     * @return The updated {@link CellStyle}
     */
    private CellStyle updateCellStyle(ExcelColumnStyle style, CellStyle defaultStyle) {

        CellStyle cellStyle = wb.createCellStyle();

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

    /**
     * Gets an {@link ExcelCellStyle} object from an {@link ExcelColumnStyle} annotation.
     * Handles both enum and class-based styles.
     *
     * @param excelColumnStyle The {@link ExcelColumnStyle} annotation
     * @return An {@link ExcelCellStyle} object
     * @throws ExcelStyleException If the style cannot be instantiated or the enum value is not
     *                             found
     */
    @SuppressWarnings("unchecked")
    private <E> ExcelCellStyle getExcelCellStyle(ExcelColumnStyle excelColumnStyle) {
        Class<? extends ExcelCellStyle> cellStyleClass = excelColumnStyle.cellStyleClass();
        //1. case of enum
        if (cellStyleClass.isEnum()) {
            try {
                Class<? extends Enum> enumClass = cellStyleClass.asSubclass(Enum.class);
                return (ExcelCellStyle) Enum.valueOf(enumClass, excelColumnStyle.enumName());
            } catch (IllegalArgumentException e) {
                throw new ExcelStyleException(
                    String.format(
                        "Failed to convert Enum constant to ExcelCellStyle instance : "
                            + "Enum value '%s' not found in style class '%s'.",
                        excelColumnStyle.enumName(), cellStyleClass.getName()), type.getName(), e);
            }
        }

        //2. case of class
        try {
            return cellStyleClass.getDeclaredConstructor().newInstance();
        } catch (NoSuchMethodException | IllegalAccessException |
                 InstantiationException | InvocationTargetException e
        ) {
            throw new ExcelStyleException(
                String.format("Failed to instantiate cellStyle class of '%s'.",
                    cellStyleClass.getName()), type.getName(), e);
        }
    }
}
