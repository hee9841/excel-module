package io.github.hee9841.excel.style.configurer;

import io.github.hee9841.excel.style.align.ExcelAlign;
import io.github.hee9841.excel.style.align.NoExcelAlign;
import io.github.hee9841.excel.style.border.ExcelBorder;
import io.github.hee9841.excel.style.border.NoExcelBorder;
import io.github.hee9841.excel.style.color.ExcelColor;
import io.github.hee9841.excel.style.color.NoExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * A configurer class that aggregates different style components (color, alignment, border)
 * and applies them to an Excel cell style.
 *
 * <p>This class follows the builder pattern, allowing clients to set individual style
 * components separately. Each component defaults to a no-option implementation if not
 * explicitly set.</p>
 *
 * <p>The configurer acts as a central point for applying multiple style components
 * to an Apache POI CellStyle in a consistent manner.</p>
 *
 * @see io.github.hee9841.excel.style.CustomExcelCellStyle
 * @see io.github.hee9841.excel.style.color.ExcelColor
 * @see io.github.hee9841.excel.style.align.ExcelAlign
 * @see io.github.hee9841.excel.style.border.ExcelBorder
 */
public class ExcelCellStyleConfigurer {

    private ExcelColor excelColor = new NoExcelColor();
    private ExcelAlign excelAlign = new NoExcelAlign();
    private ExcelBorder excelBorder = new NoExcelBorder();


    public ExcelCellStyleConfigurer() {
    }

    /**
     * Sets the color component for this configurer.
     *
     * @param excelColor the color component to use
     */
    public void excelColor(ExcelColor excelColor) {
        this.excelColor = excelColor;
    }

    /**
     * Sets the border component for this configurer.
     *
     * @param excelBorder the border component to use
     */
    public void excelBorder(ExcelBorder excelBorder) {
        this.excelBorder = excelBorder;
    }

    /**
     * Sets the alignment component for this configurer.
     *
     * @param excelAlign the alignment component to use
     */
    public void excelAlign(ExcelAlign excelAlign) {
        this.excelAlign = excelAlign;
    }

    /**
     * Applies all configured style components to the provided cell style.
     * The method applies color, alignment, and border settings in sequence.
     *
     * @param cellStyle the cell style to which the configurations will be applied
     */
    public void configure(CellStyle cellStyle) {
        excelColor.applyBackground(cellStyle);
        excelAlign.applyAlign(cellStyle);
        excelBorder.applyAllBorder(cellStyle);
    }

}
