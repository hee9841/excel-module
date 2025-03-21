package io.github.hee9841.excel.style.color;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Interface defining color behavior for Excel cell backgrounds.
 * Implementations of this interface provide different strategies for
 * applying background colors to Excel cell styles.
 * 
 * @see io.github.hee9841.excel.style.color.IndexedExcelColor
 * @see io.github.hee9841.excel.style.color.RgbExcelColor
 */
public interface ExcelColor {
    /**
     * Applies a background color to the provided cell style.
     * The specific color applied depends on the implementation.
     *
     * @param cellStyle The cell style to which the background color will be applied
     */
    void applyBackground(CellStyle cellStyle);
}
