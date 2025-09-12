package io.github.hee9841.excel.strategy;

/**
 * Strategy enum that determines how column indices are assigned in Excel sheets.
 * <p>
 * This enum defines strategies for determining the order and position of columns
 * when exporting data to Excel. It controls whether column indices are determined
 * automatically based on field declaration order or explicitly defined by the user.
 * <p>
 * The available strategies are:
 * <ul>
 *   <li>{@code FIELD_ORDER}: Column indices are assigned based on the order that fields
 *       are declared in the class. This is the simplest approach and requires no
 *       additional configuration.</li>
 *   <li>{@code USER_DEFINED}: Column indices are explicitly defined by the user through
 *       the {@link io.github.hee9841.excel.annotation.ExcelColumn#columnIndex()} parameter.
 *       This gives precise control over column positioning.</li>
 * </ul>
 * <p>
 * This strategy is typically configured at the class level using the
 * {@link io.github.hee9841.excel.annotation.Excel#columnIndexStrategy()} annotation parameter.
 *
 * @see io.github.hee9841.excel.annotation.Excel#columnIndexStrategy()
 * @see io.github.hee9841.excel.annotation.ExcelColumn#columnIndex()
 * @see io.github.hee9841.excel.core.meta.ColumnInfoMapper
 */
public enum ColumnIndexStrategy {
    /**
     * Strategy that assigns column indices based on the order of field declarations.
     * <p>
     * When this strategy is used, the
     * {@link io.github.hee9841.excel.annotation.ExcelColumn#columnIndex()}
     * parameter is ignored, and columns appear in the Excel sheet in the same order
     * that fields are declared in the Java class.
     */
    FIELD_ORDER,

    /**
     * Strategy that assigns column indices based on explicit user definition.
     * <p>
     * When this strategy is used, the
     * {@link io.github.hee9841.excel.annotation.ExcelColumn#columnIndex()}
     * parameter must be specified for each column and determines its column index in the Excel
     * sheet.
     * If two columns have the same index, an exception will be thrown.
     */
    USER_DEFINED,
    ;

    public boolean isUserDefined() {
        return this == USER_DEFINED;
    }

    public boolean isFieldOrder() {
        return this == FIELD_ORDER;
    }
}
