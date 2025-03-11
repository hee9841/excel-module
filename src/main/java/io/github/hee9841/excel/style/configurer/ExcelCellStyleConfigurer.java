package io.github.hee9841.excel.style.configurer;

import io.github.hee9841.excel.style.align.ExcelAlign;
import io.github.hee9841.excel.style.align.NoExcelAlign;
import io.github.hee9841.excel.style.border.ExcelBorder;
import io.github.hee9841.excel.style.border.NoExcelBorder;
import io.github.hee9841.excel.style.color.ExcelColor;
import io.github.hee9841.excel.style.color.NoExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelCellStyleConfigurer {

    private ExcelColor excelColor = new NoExcelColor();
    private ExcelAlign excelAlign = new NoExcelAlign();
    private ExcelBorder excelBorder = new NoExcelBorder();

    public ExcelCellStyleConfigurer() {
    }

    public void excelColor(ExcelColor excelColor) {
        this.excelColor = excelColor;
    }

    public void excelBorder(ExcelBorder excelBorder) {
        this.excelBorder = excelBorder;
    }

    public void excelAlign(ExcelAlign excelAlign) {
        this.excelAlign = excelAlign;
    }

    public void configure(CellStyle cellStyle) {
        excelColor.applyBackground(cellStyle);
        excelAlign.applyAlign(cellStyle);
        excelBorder.applyAllBorder(cellStyle);
    }

}
