package io.github.hee9841.excel.example.style;

import io.github.hee9841.excel.style.CustomExcelCellStyle;
import io.github.hee9841.excel.style.align.DefaultExcelAlign;
import io.github.hee9841.excel.style.border.BorderStyle;
import io.github.hee9841.excel.style.border.DefaultExcelBorder;
import io.github.hee9841.excel.style.color.IndexedColors;
import io.github.hee9841.excel.style.color.IndexedExcelColor;
import io.github.hee9841.excel.style.configurer.ExcelCellStyleConfigurer;

public class ExcelCustomStyleExample extends CustomExcelCellStyle {

    public ExcelCustomStyleExample() {
        super();
    }

    @Override
    public void configure(ExcelCellStyleConfigurer configurer) {
        //Set Color
        //1. Use IndexedColors
        configurer.excelColor(IndexedExcelColor.of(IndexedColors.GREY_25_PERCENT));
        // 2. Use rgb
        //configurer.excelColor(RgbExcelColor.rgb(100, 100, 100));

        // Set Align
        // 1. Use default align
        configurer.excelAlign(DefaultExcelAlign.GENERAL_CENTER);
        // 2. Use custom align
        // configurer.excelAlign(CustomExcelAlign.of(HORIZONTAL_GENERAL,VERTICAL_CENTER));
        // configurer.excelAlign(CustomExcelAlign.from(HORIZONTAL_GENERAL));
        // configurer.excelAlign(CustomExcelAlign.from(VERTICAL_CENTER));

        // Set Border
        //1. Use all
        configurer.excelBorder(DefaultExcelBorder.all(BorderStyle.THIN));
        //2. Use builder
//        configurer.excelBorder(DefaultExcelBorder.builder().top(BorderStyle.THIN).build());
    }
}
