package io.github.hee9841.excel.style.align;

import io.github.hee9841.excel.style.align.alignment.HorizontalAlignment;
import io.github.hee9841.excel.style.align.alignment.VerticalAlignment;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Implementation of ExcelAlign that allows for custom configuration of 
 * horizontal and vertical alignment settings for Excel cells.
 * This class provides flexibility by allowing either or both alignment types to be specified.
 */
public class CustomExcelAlign implements ExcelAlign {

    private final HorizontalAlignment horizontalAlignment;
    private final VerticalAlignment verticalAlignment;

    private CustomExcelAlign(HorizontalAlignment horizontalAlignment,
        VerticalAlignment verticalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        this.verticalAlignment = verticalAlignment;
    }

    /**
     * Creates a CustomExcelAlign instance with both horizontal and vertical alignment settings.
     *
     * @param horizontalAlignment the horizontal alignment to apply
     * @param verticalAlignment the vertical alignment to apply
     * @return a new CustomExcelAlign instance
     */
    public static CustomExcelAlign of(HorizontalAlignment horizontalAlignment,
        VerticalAlignment verticalAlignment) {
        return new CustomExcelAlign(
            horizontalAlignment,
            verticalAlignment
        );
    }

    /**
     * Creates a CustomExcelAlign instance with only horizontal alignment specified.
     * Vertical alignment will remain null.
     *
     * @param horizontalAlignment the horizontal alignment to apply
     * @return a new CustomExcelAlign instance
     */
    public static CustomExcelAlign from(HorizontalAlignment horizontalAlignment) {
        return new CustomExcelAlign(
            horizontalAlignment,
            null
        );
    }

    /**
     * Creates a CustomExcelAlign instance with only vertical alignment specified.
     * Horizontal alignment will remain null.
     *
     * @param verticalAlignment the vertical alignment to apply
     * @return a new CustomExcelAlign instance
     */
    public static CustomExcelAlign from(VerticalAlignment verticalAlignment) {
        return new CustomExcelAlign(
            null,
            verticalAlignment
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
        if (horizontalAlignment != null) {
            cellStyle.setAlignment(horizontalAlignment.getAlign());
        }
        if (verticalAlignment != null) {
            cellStyle.setVerticalAlignment(verticalAlignment.getAlign());
        }
    }
}
