package io.github.hee9841.excel.style.align;

import io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment;
import io.github.hee9841.excel.style.align.alignment.ExcelVerticalAlignment;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Implementation of ExcelAlign that allows for custom configuration of
 * horizontal and vertical alignment settings for Excel cells.
 * This class provides flexibility by allowing either or both alignment types to be specified.
 */
public class CustomExcelAlign implements ExcelAlign {

    private final ExcelHorizontalAlignment excelHorizontalAlignment;
    private final ExcelVerticalAlignment excelVerticalAlignment;

    private CustomExcelAlign(ExcelHorizontalAlignment excelHorizontalAlignment,
        ExcelVerticalAlignment excelVerticalAlignment) {
        this.excelHorizontalAlignment = excelHorizontalAlignment;
        this.excelVerticalAlignment = excelVerticalAlignment;
    }

    /**
     * Creates a CustomExcelAlign instance with both horizontal and vertical alignment settings.
     *
     * @param excelHorizontalAlignment the horizontal alignment to apply
     * @param excelVerticalAlignment   the vertical alignment to apply
     * @return a new CustomExcelAlign instance
     */
    public static CustomExcelAlign of(ExcelHorizontalAlignment excelHorizontalAlignment,
        ExcelVerticalAlignment excelVerticalAlignment) {
        return new CustomExcelAlign(
            excelHorizontalAlignment,
            excelVerticalAlignment
        );
    }

    /**
     * Creates a CustomExcelAlign instance with only horizontal alignment specified.
     * Vertical alignment will be set to VERTICAL_CENTER by default.
     *
     * @param excelHorizontalAlignment the horizontal alignment to apply
     * @return a new CustomExcelAlign instance
     */
    public static CustomExcelAlign from(ExcelHorizontalAlignment excelHorizontalAlignment) {
        return new CustomExcelAlign(
            excelHorizontalAlignment,
            ExcelVerticalAlignment.VERTICAL_CENTER
        );
    }

    /**
     * Creates a CustomExcelAlign instance with only vertical alignment specified.
     * Horizontal alignment will remain null.
     *
     * @param excelVerticalAlignment the vertical alignment to apply
     * @return a new CustomExcelAlign instance
     */
    public static CustomExcelAlign from(ExcelVerticalAlignment excelVerticalAlignment) {
        return new CustomExcelAlign(
            null,
            excelVerticalAlignment
        );
    }


    /**
     * Applies horizontal and/or vertical alignment to the given cell style.
     * If either alignment is null, that particular alignment type will not be modified.
     *
     * @param cellStyle the cell style to which alignment will be applied
     */
    @Override
    public void applyAlign(CellStyle cellStyle) {
        if (excelHorizontalAlignment != null) {
            cellStyle.setAlignment(excelHorizontalAlignment.getAlign());
        }
        if (excelVerticalAlignment != null) {
            cellStyle.setVerticalAlignment(excelVerticalAlignment.getAlign());
        }
    }
}
