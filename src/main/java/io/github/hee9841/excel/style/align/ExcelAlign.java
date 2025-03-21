package io.github.hee9841.excel.style.align;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Interface defining alignment behavior for Excel cells.
 * Implementations of this interface provide different strategies for
 * applying horizontal and vertical alignment to Excel cell styles.
 * 
 * @see io.github.hee9841.excel.style.align.DefaultExcelAlign
 * @see io.github.hee9841.excel.style.align.CustomExcelAlign
 */
public interface ExcelAlign {

    /**
     * Applies alignment settings to the provided cell style.
     *
     * @param cellStyle The cell style to which alignment will be applied
     */
    void applyAlign(CellStyle cellStyle);
}
