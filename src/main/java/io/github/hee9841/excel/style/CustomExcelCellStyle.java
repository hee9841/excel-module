package io.github.hee9841.excel.style;

import io.github.hee9841.excel.style.configurer.ExcelCellStyleConfigurer;
import org.apache.poi.ss.usermodel.CellStyle;

public abstract class CustomExcelCellStyle implements ExcelCellStyle {

    private final ExcelCellStyleConfigurer configurer = new ExcelCellStyleConfigurer();

    public CustomExcelCellStyle() {
        configure(configurer);
    }

    public abstract void configure(ExcelCellStyleConfigurer configurer);

    @Override
    public void apply(CellStyle cellStyle) {
        configurer.configure(cellStyle);
    }

}
