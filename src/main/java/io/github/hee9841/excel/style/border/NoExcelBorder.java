package io.github.hee9841.excel.style.border;

import org.apache.poi.ss.usermodel.CellStyle;

public class NoExcelBorder implements ExcelBorder {

    @Override
    public void applyAllBorder(CellStyle cellStyle) {
        // do nothing.
    }
}
