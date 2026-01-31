package io.github.hee9841.excel.annotation;

import io.github.hee9841.excel.mode.CellTypeMode;
import io.github.hee9841.excel.mode.ColumnIndexMode;
import io.github.hee9841.excel.mode.DataFormatMode;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation used to configure Excel export settings for a class.
 * This annotation should be applied at the class level to specify how the class
 * should be processed when converting to/from Excel format.
 *
 * <p>The class annotated with @Excel must:</p>
 * <ul>
 *   <li>Have fields annotated with @ExcelColumn to map to Excel columns</li>
 *   <li>Not be an abstract class or interface</li>
 * </ul>
 *
 * <p>Example usage:</p>
 * <pre><code>
 * {@literal @}Excel(
 *     cellTypeMode = CellTypeMode.AUTO,
 *     columnIndexMode = ColumnIndexMode.USER_DEFINED
 * )
 * public class UserData {
 *     {@literal @}ExcelColumn(headerName = "User ID", columnIndex = 0)
 *     private Long id;
 *
 *     {@literal @}ExcelColumn(headerName = "User Name", columnIndex = 1)
 *     private String name;
 *
 *     {@literal @}ExcelColumn(headerName = "Registration Date", columnIndex = 2,
 *                 format = "yyyy-MM-dd")
 *     private LocalDate registrationDate;
 *
 *     // getters and setters...
 * }
 *
 * // Usage example:
 * List{@literal <}UserData{@literal >} users = getUserData();
 * SXSSFExporter{@literal <}UserData{@literal >} exporter = SXSSFExporter.builder(UserData.class, users)
 *     .sheetName("Users")
 *     .build();
 * exporter.write(new FileOutputStream("users.xlsx"));
 * </code></pre>
 *
 * @see ExcelColumn
 * @see ExcelColumnStyle
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

    /**
     * Specifies the mode for determining column indices in the Excel sheet.
     * Default is FIELD_ORDER which uses the order of fields in the class.
     *
     * @return the column index mode to use
     */
    ColumnIndexMode columnIndexMode() default ColumnIndexMode.FIELD_ORDER;

    /**
     * Specifies the mode for determining cell types in the Excel sheet.
     * Default is NONE which means no specific cell type mode is applied.
     *
     * @return the cell type mode to use
     */
    CellTypeMode cellTypeMode() default CellTypeMode.NONE;

    /**
     * Specifies the mode for formatting data in the Excel sheet.
     * Default is NONE which means no specific data format mode is applied.
     *
     * @return the data format mode to use
     */
    DataFormatMode dataFormatMode() default DataFormatMode.NONE;

    /**
     * Specifies the default style to be applied to header cells.
     * This style will be used for all columns unless overridden by individual column styles.
     *
     * @return the default header cell style
     */
    ExcelColumnStyle defaultHeaderStyle() default @ExcelColumnStyle;

    /**
     * Specifies the default style to be applied to body cells.
     * This style will be used for all columns unless overridden by individual column styles.
     *
     * @return the default body cell style
     */
    ExcelColumnStyle defaultBodyStyle() default @ExcelColumnStyle;
}
