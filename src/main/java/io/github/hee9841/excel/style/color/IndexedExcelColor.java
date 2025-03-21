package io.github.hee9841.excel.style.color;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;

/**
 * Implementation of ExcelColor that uses Excel's built-in indexed color palette
 * to apply background colors to cell styles.
 * 
 * <p>This class works with both HSSF (XLS) and XSSF (XLSX) formats, as it uses
 * the standard indexed color palette that is compatible with all Excel versions.</p>
 * 
 * @see io.github.hee9841.excel.style.color.IndexedColors
 */
public class IndexedExcelColor implements ExcelColor{

    private final short index;

    private IndexedExcelColor(short index) {
        this.index = index;
    }

    
    public static IndexedExcelColor of(IndexedColors indexedColors) {
        return new IndexedExcelColor(indexedColors.getIndex());
    }

    /**
     * Applies the indexed color as a solid background to the provided cell style.
     * Sets both the fill foreground color and the fill pattern.
     *
     * @param cellStyle the cell style to which the background color will be applied
     */
    @Override
    public void applyBackground(CellStyle cellStyle) {
        cellStyle.setFillForegroundColor(index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
}
