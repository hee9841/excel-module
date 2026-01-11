package io.github.hee9841.excel.core.meta;

import static io.github.hee9841.excel.global.AllowedTypes.*;

import io.github.hee9841.excel.annotation.ExcelColumnStyle;
import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.format.CellFormats;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.function.BiConsumer;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Enum defining the supported cell types for Excel export/import operations.
 * Each type defines how to convert Java values to Excel cell values and which Java types are
 * supported.
 * The enum also provides utility methods for type matching and cell value setting.
 *
 * @see ColumnInfoMapper
 * @see ColumnInfo
 * @see io.github.hee9841.excel.annotation.Excel
 * @see io.github.hee9841.excel.annotation.ExcelColumn
 * @see ExcelColumnStyle
 * @see CellFormats
 */
public enum ColumnDataType {
    /**
     * Automatically determine the cell type based on the field type
     */
    AUTO,

    /**
     * No specific cell type (default)
     */
    _NONE,

    /**
     * Numeric cell type for various number formats
     */
    NUMBER(
        typedSetter(Number.class, (cell, o) -> cell.setCellValue(o.doubleValue())),
        NUMBER_TYPES,
        CellFormats._NONE,
        true
    ),

    /**
     * Boolean cell type
     */
    BOOLEAN(
        typedSetter(Boolean.class, Cell::setCellValue),
        BOOLEAN_TYPES,
        CellFormats._NONE,
        true
    ),

    /**
     * String cell type for text values
     */
    STRING(
        typedSetter(Object.class, (cell, o) -> cell.setCellValue(String.valueOf(o))),
        STRING_TYPES,
        CellFormats._NONE,
        true
    ),

    /**
     * Enum cell type - uses toString() to get the value
     */
    ENUM(
        typedSetter(Enum.class, (cell, o) -> cell.setCellValue(o != null ? o.toString() : "")),
        ENUM_TYPES,
        CellFormats._NONE,
        true
    ),

    /**
     * Formula cell type - value is treated as an Excel formula
     */
    FORMULA(
        typedSetter(String.class, Cell::setCellFormula),
        FORMULA_TYPES,
        CellFormats._NONE,
        false
    ),

    /**
     * Date And Time cell type
     */
    DATE(
        typedSetter(Date.class, Cell::setCellValue),
        DATE_TYPES,
        CellFormats.DEFAULT_DATE_FORMAT,
        true
    ),
    LOCAL_DATE(
        typedSetter(LocalDate.class, Cell::setCellValue),
        LOCAL_DATE_TYPES,
        CellFormats.DEFAULT_DATE_FORMAT,
        true
    ),
    LOCAL_DATE_TIME(
        typedSetter(LocalDateTime.class, Cell::setCellValue),
        LOCAL_DATE_TIME_TYPES,
        CellFormats.DEFAULT_DATE_TIME_FORMAT,
        true
    );

    /**
     * Function to set a cell's value based on the given object.
     */
    private final TypedCellValueSetter<?> cellValueSetter;
    /**
     * List of Java types allowed for this cell type
     */
    private final List<Class<?>> allowedTypes;
    /**
     * Default data format pattern by this cell type
     */
    private final String dataFormatPattern;
    /**
     * Whether this cell type has high priority when has same allowed types
     */
    private final boolean hasHighPriority;


    ColumnDataType(
        TypedCellValueSetter<?> cellValueSetter,
        List<Class<?>> allowedTypes,
        String dataFormatPattern,
        boolean hasHighPriority
    ) {
        this.cellValueSetter = cellValueSetter;
        this.dataFormatPattern = dataFormatPattern;
        this.allowedTypes = allowedTypes;
        this.hasHighPriority = hasHighPriority;
    }

    /**
     * Default constructor for special {@link ColumnDataType} (AUTO and
     * {@link ColumnDataType#_NONE}).
     * Uses a default string value setter, empty allowed types list,
     * and no specific format pattern.
     */
    ColumnDataType() {
        this(
            typedSetter(Object.class, (cell, o) -> cell.setCellValue(String.valueOf(o))),
            Collections.emptyList(),
            CellFormats._NONE,
            false
        );
    }

    /**
     * Determines the appropriate {@link ColumnDataType} based on the given field type.
     * Searches through all {@link ColumnDataType} that have high priority and
     * finds the first one where the field type is assignable to one of the allowed types.
     *
     * @param fieldType The Java type to match against cell types
     * @return The matching {@link ColumnDataType}, or {@link ColumnDataType#_NONE} if no match is found
     */
    public static ColumnDataType from(Class<?> fieldType) {
        return Arrays.stream(values())
            .filter(
                cellType -> cellType.hasHighPriority &&
                    cellType.allowedTypes.stream().anyMatch(c -> c.isAssignableFrom(fieldType))
            )
            .findFirst()
            .orElse(_NONE);
    }

    /**
     * Checks if fieldType matches one of the allowedTypes in targetCellType,
     * returns targetCellType if matched, otherwise returns {@link ColumnDataType#_NONE}.
     *
     * @param fieldType            The field type
     * @param targetColumnDataType specific {@link ColumnDataType}
     * @return Returns targetCellType if matched, otherwise returns {@link ColumnDataType#_NONE}
     */
    public static ColumnDataType findMatchingCellType(Class<?> fieldType,
        ColumnDataType targetColumnDataType) {
        return targetColumnDataType.allowedTypes.stream()
            .anyMatch(c -> c.isAssignableFrom(fieldType)) ? targetColumnDataType : _NONE;
    }

    /**
     * Sets a cell's value according to this {@link ColumnDataType}.
     * If the value is null, an empty string will be set.
     *
     * @param cell  The Excel cell to set the value for
     * @param value The value to set in the cell
     * @throws ExcelException If the value cannot be set for any reason
     */
    public void setCellValueByCellType(Cell cell, Object value) {
        if (value == null) {
            _NONE.cellValueSetter.accept(cell, "");
            return;
        }
        try {
            cellValueSetter.accept(cell, value);
        } catch (Exception e) {
            throw new ExcelException("Failed to set cell value by cell type: " + e.getMessage());
        }
    }

    private static <T> TypedCellValueSetter<T> typedSetter(
        Class<T> type,
        BiConsumer<Cell, T> setter
    ) {
        return new TypedCellValueSetter<>(type, setter);
    }

    private static final class TypedCellValueSetter<T> {
        private final Class<T> type;
        private final BiConsumer<Cell, T> setter;

        private TypedCellValueSetter(Class<T> type, BiConsumer<Cell, T> setter) {
            this.type = type;
            this.setter = setter;
        }

        private void accept(Cell cell, Object value) {
            setter.accept(cell, type.cast(value));
        }
    }

    public boolean isAuto() {
        return this == AUTO;
    }


    public boolean isNone() {
        return this == _NONE;
    }

    public String getDataFormatPattern() {
        return dataFormatPattern;
    }

}

