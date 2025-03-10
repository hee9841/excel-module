package io.github.hee9841.excel.style.border;

import org.apache.poi.ss.usermodel.CellStyle;

public class DefaultExcelBorder implements ExcelBorder {

    private final BorderStyle top;
    private final BorderStyle bottom;
    private final BorderStyle left;
    private final BorderStyle right;

    DefaultExcelBorder(
        BorderStyle top,
        BorderStyle bottom,
        BorderStyle left,
        BorderStyle right) {
        this.top = top;
        this.bottom = bottom;
        this.left = left;
        this.right = right;
    }


    public static DefaultExcelBorder all(BorderStyle all) {
        return new DefaultExcelBorder(all, all, all, all);
    }

    public static DefaultExcelBorderBuilder builder() {
        return new DefaultExcelBorderBuilder();
    }


    @Override
    public void applyAllBorder(CellStyle cellStyle) {
        applyTop(cellStyle);
        applyBottom(cellStyle);
        applyLeft(cellStyle);
        applyRight(cellStyle);
    }


    private void applyTop(CellStyle cellStyle) {
        if (top == null) {
            return;
        }
        cellStyle.setBorderTop(top.getBorderStyle());
    }


    private void applyBottom(CellStyle cellStyle) {
        if (bottom == null) {
            return;
        }
        cellStyle.setBorderBottom(bottom.getBorderStyle());
    }


    private void applyLeft(CellStyle cellStyle) {
        if (left == null) {
            return;
        }
        cellStyle.setBorderLeft(left.getBorderStyle());
    }


    private void applyRight(CellStyle cellStyle) {
        if (right == null) {
            return;
        }
        cellStyle.setBorderRight(right.getBorderStyle());
    }
}
