package io.github.hee9841.excel.style.align.alignment;

import static org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM;
import static org.apache.poi.ss.usermodel.VerticalAlignment.CENTER;
import static org.apache.poi.ss.usermodel.VerticalAlignment.DISTRIBUTED;
import static org.apache.poi.ss.usermodel.VerticalAlignment.JUSTIFY;
import static org.apache.poi.ss.usermodel.VerticalAlignment.TOP;

public enum VerticalAlignment {
    VERTICAL_TOP(TOP),
    VERTICAL_CENTER(CENTER),
    VERTICAL_BOTTOM(BOTTOM),
    VERTICAL_JUSTIFY(JUSTIFY),
    VERTICAL_DISTRIBUTED(DISTRIBUTED);

    private final org.apache.poi.ss.usermodel.VerticalAlignment verticalAlign;

    VerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    public org.apache.poi.ss.usermodel.VerticalAlignment getAlign() {
        return verticalAlign;
    }
}
