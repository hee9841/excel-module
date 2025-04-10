package io.github.hee9841.excel.style.align.alignment;

import static org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER_SELECTION;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.DISTRIBUTED;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.FILL;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.JUSTIFY;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT;

public enum ExcelHorizontalAlignment {
    HORIZONTAL_GENERAL(GENERAL),
    HORIZONTAL_LEFT(LEFT),
    HORIZONTAL_CENTER(CENTER),
    HORIZONTAL_RIGHT(RIGHT),
    HORIZONTAL_FILL(FILL),
    HORIZONTAL_JUSTIFY(JUSTIFY),
    HORIZONTAL_CENTER_SELECTION(CENTER_SELECTION),
    HORIZONTAL_DISTRIBUTED(DISTRIBUTED);

    private final org.apache.poi.ss.usermodel.HorizontalAlignment horizontalAlignment;

    ExcelHorizontalAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public org.apache.poi.ss.usermodel.HorizontalAlignment getAlign() {
        return horizontalAlignment;
    }
}
