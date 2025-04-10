package io.github.hee9841.excel.annotation;

import io.github.hee9841.excel.strategy.CellTypeStrategy;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import io.github.hee9841.excel.strategy.DataFormatStrategy;
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
 *     cellTypeStrategy = CellTypeStrategy.AUTO,
 *     columnIndexStrategy = ColumnIndexStrategy.USER_DEFINED
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
 * ExcelExporter{@literal <}UserData{@literal >} exporter = ExcelExporter.builder(UserData.class, users)
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
     * Specifies the strategy for determining column indices in the Excel sheet.
     * Default is FIELD_ORDER which uses the order of fields in the class.
     *
     * @return the column index strategy to use
     */
    ColumnIndexStrategy columnIndexStrategy() default ColumnIndexStrategy.FIELD_ORDER;

    /**
     * Specifies the strategy for determining cell types in the Excel sheet.
     * Default is NONE which means no specific cell type strategy is applied.
     *
     * @return the cell type strategy to use
     */
    CellTypeStrategy cellTypeStrategy() default CellTypeStrategy.NONE;

    /**
     * Specifies the strategy for formatting data in the Excel sheet.
     * Default is NONE which means no specific data format strategy is applied.
     *
     * @return the data format strategy to use
     */
    DataFormatStrategy dataFormatStrategy() default DataFormatStrategy.NONE;

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
