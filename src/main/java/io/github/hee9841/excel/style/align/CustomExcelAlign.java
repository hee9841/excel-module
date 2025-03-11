package io.github.hee9841.excel.style.align;

import io.github.hee9841.excel.style.align.alignment.HorizontalAlignment;
import io.github.hee9841.excel.style.align.alignment.VerticalAlignment;
import org.apache.poi.ss.usermodel.CellStyle;

public class CustomExcelAlign implements ExcelAlign {

    private final HorizontalAlignment horizontalAlignment;
    private final VerticalAlignment verticalAlignment;

    private CustomExcelAlign(HorizontalAlignment horizontalAlignment,
        VerticalAlignment verticalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        this.verticalAlignment = verticalAlignment;
    }

    public static CustomExcelAlign of(HorizontalAlignment horizontalAlignment,
        VerticalAlignment verticalAlignment) {
        return new CustomExcelAlign(
            horizontalAlignment,
            verticalAlignment
        );
    }

    public static CustomExcelAlign from(HorizontalAlignment horizontalAlignment) {
        return new CustomExcelAlign(
            horizontalAlignment,
            null
        );
    }

    public static CustomExcelAlign from(VerticalAlignment verticalAlignment) {
        return new CustomExcelAlign(
            null,
            verticalAlignment
        );
    }


    @Override
    public void applyAlign(CellStyle cellStyle) {
        if (horizontalAlignment != null) {
            cellStyle.setAlignment(horizontalAlignment.getAlign());
        }
        if (verticalAlignment != null) {
            cellStyle.setVerticalAlignment(verticalAlignment.getAlign());
        }
    }
}
