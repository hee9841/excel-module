package io.github.hee9841.excel.meta;

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

public enum CellType {
    AUTO,  //auto mapping by field type
    _NONE,
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
    BOOLEAN(
        (cell, o) -> cell.setCellValue((boolean) o),
        Collections.unmodifiableList(
            Arrays.asList(Boolean.TYPE, Boolean.class)
        ),
        CellFormats._NONE,
        true
    ),
    STRING(
        (cell, o) -> cell.setCellValue(String.valueOf(o)),
        Collections.unmodifiableList(
            Arrays.asList(String.class, Character.class, char.class)
        ),
        CellFormats._NONE,
        true
    ),
    //enum의 값(name) -> toString() override
    ENUM(
        (cell, o) -> cell.setCellValue(o != null ? o.toString() : ""),
        Collections.singletonList(Enum.class),
        CellFormats._NONE,
        true
    ),
    FORMULA(
        (cell, o) -> cell.setCellFormula(String.valueOf(o)),
        Collections.singletonList(String.class),
        CellFormats._NONE,
        false
    ),
    //date
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

    private final BiConsumer<Cell, Object> cellValueSetter;
    private final List<Class<?>> allowedTypes;
    private final String dataFormatPattern;
    private final boolean hasHighPriority;


    CellType(
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

    CellType() {
        this(
            (cell, o) -> cell.setCellValue(String.valueOf(o)),
            Collections.emptyList(),
            CellFormats._NONE,
            false
        );
    }

    public static CellType from(Class<?> fieldType) {
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
     * returns targetCellType if matched, otherwise returns _NONE.
     *
     * @param fieldType      The field type
     * @param targetCellType specific cell type
     * @return Returns targetCellType if matched, otherwise returns _NONE
     */
    public static CellType findMatchingCellType(Class<?> fieldType, CellType targetCellType) {
        return targetCellType.allowedTypes.stream()
            .anyMatch(c -> c.isAssignableFrom(fieldType)) ? targetCellType : _NONE;
    }


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

