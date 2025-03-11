package io.github.hee9841.excel.style.color;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;

public class IndexedExcelColor implements ExcelColor{

    public final short index;

    private IndexedExcelColor(short index) {
        this.index = index;
    }

    public static IndexedExcelColor of(IndexedColors indexedColors) {
        return new IndexedExcelColor(indexedColors.getIndex());
    }


    @Override
    public void applyBackground(CellStyle cellStyle) {
        cellStyle.setFillForegroundColor(index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
}
