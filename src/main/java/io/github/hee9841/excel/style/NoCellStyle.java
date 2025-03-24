package io.github.hee9841.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;

public class NoCellStyle implements ExcelCellStyle{

    @Override
    public void apply(CellStyle cellStyle) {
        //do nothing.
    }
}
