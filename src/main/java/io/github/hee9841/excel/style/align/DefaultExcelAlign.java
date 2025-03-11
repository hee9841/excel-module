package io.github.hee9841.excel.style.align;

import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_CENTER;
import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_CENTER_SELECTION;
import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_DISTRIBUTED;
import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_FILL;
import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_GENERAL;
import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_JUSTIFY;
import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_LEFT;
import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_RIGHT;
import static io.github.hee9841.excel.style.align.alignment.VerticalAlignment.VERTICAL_CENTER;

import io.github.hee9841.excel.style.align.alignment.HorizontalAlignment;
import io.github.hee9841.excel.style.align.alignment.VerticalAlignment;
import org.apache.poi.ss.usermodel.CellStyle;

public enum DefaultExcelAlign implements ExcelAlign {
    GENERAL_CENTER(HORIZONTAL_GENERAL, VERTICAL_CENTER),
    LEFT_CENTER(HORIZONTAL_LEFT, VERTICAL_CENTER),
    CENTER_CENTER(HORIZONTAL_CENTER, VERTICAL_CENTER),
    RIGHT_CENTER(HORIZONTAL_RIGHT, VERTICAL_CENTER),
    FILL_CENTER(HORIZONTAL_FILL, VERTICAL_CENTER),
    JUSTIFY_CENTER(HORIZONTAL_JUSTIFY, VERTICAL_CENTER),
    CENTER_SELECTION_CENTER(HORIZONTAL_CENTER_SELECTION, VERTICAL_CENTER),
    DISTRIBUTED_CENTER(HORIZONTAL_DISTRIBUTED, VERTICAL_CENTER),
    ;

    private final HorizontalAlignment horizontalAlignment;
    private final VerticalAlignment verticalAlignment;


    DefaultExcelAlign(HorizontalAlignment horizontalAlignment,
        VerticalAlignment verticalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        this.verticalAlignment = verticalAlignment;
    }

    @Override
    public void applyAlign(CellStyle cellStyle) {
        cellStyle.setAlignment(horizontalAlignment.getAlign());
        cellStyle.setVerticalAlignment(verticalAlignment.getAlign());
    }
}
