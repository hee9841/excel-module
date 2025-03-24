package io.github.hee9841.excel.strategy;

/**
 * Strategy enum that determines how data formatting is applied to Excel cells.
 * <p>
 * This enum defines strategies for controlling how cell formatting patterns are determined and applied during
 * Excel export operations. It works in conjunction with the {@link io.github.hee9841.excel.meta.CellType}
 * to apply appropriate formatting to different types of data.
 * <p>
 * The available strategies are:
 * <ul>
 *   <li>{@code NONE}: No automatic formatting is applied. Formatting must be explicitly specified
 *       in the {@link io.github.hee9841.excel.annotation.ExcelColumn} annotation.</li>
 *   <li>{@code AUTO_BY_CELL_TYPE}: Automatic formatting is applied based on the {@link io.github.hee9841.excel.meta.CellType}, but only when the
 *       {@link io.github.hee9841.excel.annotation.ExcelColumn#columnCellType()} is explicitly set or determined by the {@link CellTypeStrategy} (not when it's NONE).
 *       If user specified the {@code format} parameter in the {@link io.github.hee9841.excel.annotation.ExcelColumn} annotation, it will take precedence over the automatic format.
 *   </li>
 * </ul>
 * <p>
 * This strategy is typically configured at the class level using the {@link io.github.hee9841.excel.annotation.Excel}
 * annotation and affects all columns unless overridden at the column level.
 * 
 * @see io.github.hee9841.excel.annotation.Excel#dataFormatStrategy()
 * @see io.github.hee9841.excel.meta.CellType
 * @see io.github.hee9841.excel.format.ExcelDataFormater
 * @see io.github.hee9841.excel.format.CellFormats
 */
public enum DataFormatStrategy {
    /**
     * No automatic formatting is applied.
     * <p>
     * With this strategy, the format pattern must be explicitly specified in the
     * {@link io.github.hee9841.excel.annotation.ExcelColumn} annotation using the
     * {@code format} parameter. If no format is specified, the default format({@link io.github.hee9841.excel.format.CellFormats#_NONE}) for
     * the data type will be used.
     */
    NONE,
    
    /**
     * Automatic formatting is applied based on the cell type.
     * <p>
     * This strategy only applies when the cell type is explicitly set (not NONE) or
     * determined by the {@link CellTypeStrategy} (when set to AUTO). It uses the format
     * pattern defined for each {@link io.github.hee9841.excel.meta.CellType} to format
     * the cell appropriately for its data type.
     * <p>
     * Note: If a format pattern is explicitly specified in the {@link io.github.hee9841.excel.annotation.ExcelColumn}
     * annotation, it will take precedence over the automatic format.
     */
    AUTO_BY_CELL_TYPE,
    ;

    public boolean isAutoByCellType() {
        return this == AUTO_BY_CELL_TYPE;
    }

    public boolean isNone() {
        return this == NONE;
    }
}
