package io.github.hee9841.excel.core.meta;

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
        (cell, o) -> cell.setCellValue(Double.parseDouble(String.valueOf(o))),
        Collections.unmodifiableList(
            Arrays.asList(
                Integer.TYPE, Double.TYPE,
                Float.TYPE, Long.TYPE,
                Short.TYPE, Byte.TYPE,
                Number.class
            )
        ),
        CellFormats._NONE,
        true
    ),

    /**
     * Boolean cell type
     */
    BOOLEAN(
        (cell, o) -> cell.setCellValue((boolean) o),
        Collections.unmodifiableList(
            Arrays.asList(Boolean.TYPE, Boolean.class)
        ),
        CellFormats._NONE,
        true
    ),

    /**
     * String cell type for text values
     */
    STRING(
        (cell, o) -> cell.setCellValue(String.valueOf(o)),
        Collections.unmodifiableList(
            Arrays.asList(String.class, Character.class, char.class)
        ),
        CellFormats._NONE,
        true
    ),

    /**
     * Enum cell type - uses toString() to get the value
     */
    ENUM(
        (cell, o) -> cell.setCellValue(o != null ? o.toString() : ""),
        Collections.singletonList(Enum.class),
        CellFormats._NONE,
        true
    ),

    /**
     * Formula cell type - value is treated as an Excel formula
     */
    FORMULA(
        (cell, o) -> cell.setCellFormula(String.valueOf(o)),
        Collections.singletonList(String.class),
        CellFormats._NONE,
        false
    ),

    /**
     * Date And Time cell type
     */
    DATE(
        (cell, o) -> cell.setCellValue((Date) o),
        Collections.unmodifiableList(
            Arrays.asList(Date.class, java.sql.Date.class)
        ),
        CellFormats.DEFAULT_DATE_FORMAT,
        true
    ),
    LOCAL_DATE(
        (cell, o) -> cell.setCellValue((LocalDate) o),
        Collections.singletonList(LocalDate.class),
        CellFormats.DEFAULT_DATE_FORMAT,
        true
    ),
    LOCAL_DATE_TIME(
        (cell, o) -> cell.setCellValue((LocalDateTime) o),
        Collections.singletonList(LocalDateTime.class),
        CellFormats.DEFAULT_DATE_TIME_FORMAT,
        true
    );

    /**
     * Function to set a cell's value based on the given object
     */
    private final BiConsumer<Cell, Object> cellValueSetter;
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
        BiConsumer<Cell, Object> cellValueSetter,
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
            (cell, o) -> cell.setCellValue(String.valueOf(o)),
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

