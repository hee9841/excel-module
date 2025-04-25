package io.github.hee9841.excel.style.align;

import static io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment.HORIZONTAL_CENTER;
import static io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment.HORIZONTAL_CENTER_SELECTION;
import static io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment.HORIZONTAL_DISTRIBUTED;
import static io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment.HORIZONTAL_FILL;
import static io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment.HORIZONTAL_GENERAL;
import static io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment.HORIZONTAL_JUSTIFY;
import static io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment.HORIZONTAL_LEFT;
import static io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment.HORIZONTAL_RIGHT;
import static io.github.hee9841.excel.style.align.alignment.ExcelVerticalAlignment.VERTICAL_BOTTOM;
import static io.github.hee9841.excel.style.align.alignment.ExcelVerticalAlignment.VERTICAL_CENTER;

import io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment;
import io.github.hee9841.excel.style.align.alignment.ExcelVerticalAlignment;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Enumeration of predefined Excel cell alignment options.
 * Provides common combinations of horizontal and vertical alignments
 * for Excel cells. Each enum constant represents a specific alignment configuration.
 *
 * <pre>
 * The enum names follow the pattern of "{HORIZONTAL}_{VERTICAL}" to clearly indicate
 * the alignment combination (e.g. LEFT_CENTER means horizontally left-aligned and
 * vertically centered).
 * </pre>
 */
public enum DefaultExcelAlign implements ExcelAlign {

    GENERAL_CENTER(HORIZONTAL_GENERAL, VERTICAL_CENTER),
    LEFT_CENTER(HORIZONTAL_LEFT, VERTICAL_CENTER),
    CENTER_CENTER(HORIZONTAL_CENTER, VERTICAL_CENTER),
    RIGHT_CENTER(HORIZONTAL_RIGHT, VERTICAL_CENTER),
    FILL_CENTER(HORIZONTAL_FILL, VERTICAL_CENTER),
    JUSTIFY_CENTER(HORIZONTAL_JUSTIFY, VERTICAL_CENTER),
    CENTER_SELECTION_CENTER(HORIZONTAL_CENTER_SELECTION, VERTICAL_CENTER),
    DISTRIBUTED_CENTER(HORIZONTAL_DISTRIBUTED, VERTICAL_CENTER),

    GENERAL_BOTTOM(HORIZONTAL_GENERAL, VERTICAL_BOTTOM),
    LEFT_BOTTOM(HORIZONTAL_LEFT, VERTICAL_BOTTOM),
    CENTER_BOTTOM(HORIZONTAL_CENTER, VERTICAL_BOTTOM),
    RIGHT_BOTTOM(HORIZONTAL_RIGHT, VERTICAL_BOTTOM),
    FILL_BOTTOM(HORIZONTAL_FILL, VERTICAL_BOTTOM),
    JUSTIFY_BOTTOM(HORIZONTAL_JUSTIFY, VERTICAL_BOTTOM),
    CENTER_SELECTION_BOTTOM(HORIZONTAL_CENTER_SELECTION, VERTICAL_BOTTOM),
    DISTRIBUTED_BOTTOM(HORIZONTAL_DISTRIBUTED, VERTICAL_BOTTOM),
    ;

    private final ExcelHorizontalAlignment excelHorizontalAlignment;
    private final ExcelVerticalAlignment excelVerticalAlignment;


    DefaultExcelAlign(ExcelHorizontalAlignment excelHorizontalAlignment,
        ExcelVerticalAlignment excelVerticalAlignment) {
        this.excelHorizontalAlignment = excelHorizontalAlignment;
        this.excelVerticalAlignment = excelVerticalAlignment;
    }

    /**
     * Applies both horizontal and vertical alignment to the given cell style.
     * Unlike CustomExcelAlign, this method always sets both alignment types.
     *
     * @param cellStyle the cell style to which alignment will be applied
     */
    @Override
    public void applyAlign(CellStyle cellStyle) {
        cellStyle.setAlignment(excelHorizontalAlignment.getAlign());
        cellStyle.setVerticalAlignment(excelVerticalAlignment.getAlign());
    }
}
