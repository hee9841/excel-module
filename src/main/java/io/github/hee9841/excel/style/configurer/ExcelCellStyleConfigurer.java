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


    public ExcelCellStyleConfigurer excelColor(ExcelColor excelColor) {
        this.excelColor = excelColor;
        return this;
    }

    public ExcelCellStyleConfigurer excelBorder(ExcelBorder excelBorder) {
        this.excelBorder = excelBorder;
        return this;
    }

    public ExcelCellStyleConfigurer excelAlign(ExcelAlign excelAlign) {
        this.excelAlign = excelAlign;
        return this;
    }

    public void configure(CellStyle cellStyle) {
        excelColor.applyBackground(cellStyle);
        excelBorder.applyAllBorder(cellStyle);
        excelAlign.applyAlign(cellStyle);
    }

}
