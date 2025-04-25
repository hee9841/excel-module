package io.github.hee9841.excel.strategy;

import io.github.hee9841.excel.meta.ColumnDataType;

/**
 * Strategy enum that determines how cell types are assigned in Excel sheets.
 * <p>
 * This enum defines strategies for determining the data type and format of cells
 * when exporting data to Excel. It works in conjunction with the
 * {@link ColumnDataType} enum to control how different Java types
 * are represented in Excel.
 * </p>
 * The available strategies are:
 * <ul>
 *   <li>{@code NONE}: Cell types are not applied.</li>
 *   <li>{@code AUTO}: Cell types are automatically determined based on the Java type of each field.</li>
 * </ul>
 * <p>
 * This strategy is typically configured at the class level using the
 * {@link io.github.hee9841.excel.annotation.Excel#cellTypeStrategy()} annotation parameter and
 * can be overridden at the column level using
 * {@link io.github.hee9841.excel.annotation.ExcelColumn#columnCellType()}.
 * </p>
 * When used with {@link DataFormatStrategy#AUTO_BY_CELL_TYPE}, it also affects how data formatting
 * is applied to cells.
 *
 * @see io.github.hee9841.excel.annotation.Excel#cellTypeStrategy()
 * @see io.github.hee9841.excel.annotation.ExcelColumn#columnCellType()
 * @see ColumnDataType
 * @see io.github.hee9841.excel.meta.ColumnInfoMapper
 * @see DataFormatStrategy
 */
public enum CellTypeStrategy {
    /**
     * Strategy that does not apply any cell type.
     * <p>
     * With this strategy, you can explicitly specify the cell type for each column
     * using the {@link io.github.hee9841.excel.annotation.ExcelColumn#columnCellType()} parameter
     * or if not specified, cell type will be {@link ColumnDataType#_NONE}.
     * </p>
     */
    NONE,

    /**
     * Strategy that automatically determines cell types based on field types.
     * <p>
     * With this strategy, the cell type is automatically determined based on the Java type
     * of each field. This simplifies configuration but may not always choose the optimal
     * representation, especially for complex types or when specific formatting is required.
     * </p>
     * <p>
     * The automatic cell type determination is performed by the
     * {@link io.github.hee9841.excel.meta.ColumnInfoMapper} class.
     * If you want to specify the cell type for a specific column, you can use the
     * {@link io.github.hee9841.excel.annotation.ExcelColumn#columnCellType()} parameter.
     * </p>
     */
    AUTO,
    ;

    public boolean isAuto() {
        return this == AUTO;
    }

    public boolean isNone() {
        return this == NONE;
    }
}
