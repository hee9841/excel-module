package io.github.hee9841.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;

public interface ExcelCellStyle {
    void apply(CellStyle cellStyle);
}
