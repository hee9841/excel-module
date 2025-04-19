package io.github.hee9841.excel.annotation;

import io.github.hee9841.excel.style.ExcelCellStyle;
import io.github.hee9841.excel.style.NoCellStyle;
import java.lang.annotation.Target;

/**
 * Annotation used to specify the styling for Excel cells.
 * This annotation is used within @Excel and @ExcelColumn annotations to define
 * the visual appearance of cells in the Excel sheet.
 *
 * @see Excel
 * @see ExcelColumn
 * @see ExcelCellStyle
 */
@Target({})
public @interface ExcelColumnStyle {

    /**
     * Specifies the class that implements the cell style.
     * This class must either:
     * <ul>
     *   <li>Implement the ExcelCellStyle interface directly</li>
     *   <li>Be an enum class that implements the ExcelCellStyle interface</li>
     *   <li>Extend the CustomExcelCellStyle abstract class</li>
     * </ul>
     * Default is NoCellStyle which means no specific styling is applied.
     *
     * @return the class implementing the cell style
     */
    Class<? extends ExcelCellStyle> cellStyleClass() default NoCellStyle.class;

    /**
     * Specifies the enum constant name to be used for cell styling.
     * This is only used when {@link #cellStyleClass()} is an enum class
     * that implements ExcelCellStyle.
     * The enum constant with this name will determine the cell style.
     * Default is an empty string which means no enum-based styling is applied.
     *
     * @return the name of the enum constant to use for cell styling
     */
    String enumName() default "";
}
