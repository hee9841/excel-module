package io.github.hee9841.excel.style.color;

import org.apache.poi.ss.usermodel.CellStyle;

public class NoExcelColor implements ExcelColor {

    @Override
    public void applyBackground(CellStyle cellStyle) {
        //do nothing
    }
}
