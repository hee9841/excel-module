package io.github.hee9841.excel.style.border;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Interface defining border behavior for Excel cells.
 * Implementations of this interface provide different strategies for
 * applying border styles to Excel cell styles.
 *
 * @see io.github.hee9841.excel.style.border.DefaultExcelBorder
 */
public interface ExcelBorder {

    /**
     * Applies border settings to all sides of a cell style.
     * The specific border styles applied to depend on the implementation.
     *
     * @param cellStyle The cell style to which borders will be applied
     */
    void applyAllBorder(CellStyle cellStyle);

}
