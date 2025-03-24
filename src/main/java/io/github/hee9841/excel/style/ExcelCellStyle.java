package io.github.hee9841.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Interface defining the core cell styling behavior for Excel cells.
 * This is the base contract for all cell styling approaches in the library.
 * 
 * <p>Implementations of this interface provide different approaches to
 * configuring and applying styles to Excel cell styles through Apache POI.</p>
 * 
 * @see io.github.hee9841.excel.style.CustomExcelCellStyle
 */
public interface ExcelCellStyle {
    /**
     * Applies the defined styling to the provided cell style.
     * 
     * @param cellStyle the Apache POI cell style to which styles will be applied
     */
    void apply(CellStyle cellStyle);
}
