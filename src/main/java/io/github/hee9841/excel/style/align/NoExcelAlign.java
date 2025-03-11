package io.github.hee9841.excel.style.align;

import org.apache.poi.ss.usermodel.CellStyle;

public class NoExcelAlign implements ExcelAlign {

    @Override
    public void applyAlign(CellStyle cellStyle) {
        //do nothing
    }
}
