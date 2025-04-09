package io.github.hee9841.excel.example.style;

import io.github.hee9841.excel.style.ExcelCellStyle;
import io.github.hee9841.excel.style.align.DefaultExcelAlign;
import io.github.hee9841.excel.style.align.ExcelAlign;
import io.github.hee9841.excel.style.align.NoExcelAlign;
import io.github.hee9841.excel.style.border.BorderStyle;
import io.github.hee9841.excel.style.border.DefaultExcelBorder;
import io.github.hee9841.excel.style.border.ExcelBorder;
import io.github.hee9841.excel.style.border.NoExcelBorder;
import io.github.hee9841.excel.style.color.ColorPalette;
import io.github.hee9841.excel.style.color.ExcelColor;
import io.github.hee9841.excel.style.color.NoExcelColor;
import io.github.hee9841.excel.style.color.PaletteExcelColor;
import io.github.hee9841.excel.style.color.RgbExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;

public enum EnumCellStyleExample implements ExcelCellStyle {
    NO_STYLE(
        new NoExcelColor(),
        new NoExcelAlign(),
        new NoExcelBorder()
    ),
    GREY_25_PERCENT_CENTER_CENTER_ALL_BORDER_THICK(
        PaletteExcelColor.of(ColorPalette.GREY_25_PERCENT),
        DefaultExcelAlign.CENTER_CENTER,
        DefaultExcelBorder.all(BorderStyle.THICK)
    ),
    GREEN_PASTEL(
        RgbExcelColor.rgb(198, 219, 218),
        new NoExcelAlign(),
        new NoExcelBorder()
    ),
    RED_CENTER_CENTER_ALL_BORDER_THICK(
        PaletteExcelColor.of(ColorPalette.RED),
        DefaultExcelAlign.CENTER_CENTER,
        DefaultExcelBorder.all(BorderStyle.THICK)
    ),
    ALL_CENTER_ALL_BORDER_THICK(
        new NoExcelColor(),
        DefaultExcelAlign.CENTER_CENTER,
        DefaultExcelBorder.all(BorderStyle.THICK)
    ),
    ;

    private final ExcelColor backgroundColor;
    private final ExcelAlign align;
    private final ExcelBorder border;

    EnumCellStyleExample(ExcelColor backgroundColor, ExcelAlign align, ExcelBorder border) {
        this.backgroundColor = backgroundColor;
        this.align = align;
        this.border = border;
    }

    @Override
    public void apply(CellStyle cellStyle) {
        backgroundColor.applyBackground(cellStyle);
        align.applyAlign(cellStyle);
        border.applyAllBorder(cellStyle);
    }
}
