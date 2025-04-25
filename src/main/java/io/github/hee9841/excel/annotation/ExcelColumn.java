package io.github.hee9841.excel.annotation;

import io.github.hee9841.excel.format.CellFormats;
import io.github.hee9841.excel.meta.ColumnDataType;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation used to configure how a field should be mapped to an Excel column.
 * This annotation should be applied to fields within a class annotated with @Excel
 * to specify the column's properties in the Excel sheet.
 *
 * @see Excel
 * @see ExcelColumnStyle
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    /**
     * The name to be displayed in the header row of the Excel sheet.
     * This is a required field and cannot be empty.
     *
     * @return the header name for the column
     */
    String headerName();

    /**
     * The index of the column in the Excel sheet (starting from 0).
     * If not specified (default -1), the column index will be determined by the
     * {@link io.github.hee9841.excel.annotation.Excel#columnIndexStrategy()} specified in the @Excel annotation.
     * When {@link io.github.hee9841.excel.annotation.Excel#columnIndexStrategy()} is
     * {@link io.github.hee9841.excel.strategy.ColumnIndexStrategy#USER_DEFINED},
     * this value must be specified.
     *
     * @return the column index (starting from 0), or -1 if not specified
     */
    int columnIndex() default -1;

    /**
     * Specifies the cell type for this column.
     * This determines how data should be interpreted and formatted in the Excel sheet.
     * <p>
     * If specified (not _NONE), this value takes precedence over
     * the {@link io.github.hee9841.excel.annotation.Excel#cellTypeStrategy()}
     * is defined in the @Excel annotation. For example,
     * if {@link io.github.hee9841.excel.annotation.Excel#cellTypeStrategy()}
     * {@link io.github.hee9841.excel.strategy.CellTypeStrategy#AUTO} but
     * this columnCellType is specified, the specified type will be used.
     *
     * @return the cell type for the column
     */
    ColumnDataType columnCellType() default ColumnDataType._NONE;

    /**
     * Specifies the format pattern for the column's data.
     * This is used to format the data when writing to Excel.
     * The format pattern can be either:
     * <ul>
     *   <li>A predefined format from {@link CellFormats}</li>
     *   <li>A custom format pattern string specified by the user</li>
     * </ul>
     * Default is {@link CellFormats#_NONE} which means no specific format is applied.
     *
     * @return the format pattern for the column
     */
    String format() default CellFormats._NONE;

    /**
     * Specifies the style to be applied to the header cell of this column.
     * If not specified, the defaultHeaderStyle from the @Excel annotation will be used.
     *
     * @return the style for the header cell
     */
    ExcelColumnStyle headerStyle() default @ExcelColumnStyle;

    /**
     * Specifies the style to be applied to the body cells of this column.
     * If not specified, the defaultBodyStyle from the @Excel annotation will be used.
     *
     * @return the style for the body cells
     */
    ExcelColumnStyle bodyStyle() default @ExcelColumnStyle;
}
