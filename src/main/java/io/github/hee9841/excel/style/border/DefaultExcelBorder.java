package io.github.hee9841.excel.style.border;

import org.apache.poi.ss.usermodel.CellStyle;

public class DefaultExcelBorder implements ExcelBorder {

    private final BorderStyle top;
    private final BorderStyle bottom;
    private final BorderStyle left;
    private final BorderStyle right;

    public static DefaultExcelBorder all(BorderStyle all) {
        return new DefaultExcelBorder(all, all, all, all);
    }

    public static Builder builder() {
        return new Builder();
    }


    private DefaultExcelBorder(
        BorderStyle top,
        BorderStyle bottom,
        BorderStyle left,
        BorderStyle right) {
        this.top = top;
        this.bottom = bottom;
        this.left = left;
        this.right = right;
    }

    private DefaultExcelBorder(Builder builder) {
        this.top = builder.top;
        this.bottom = builder.bottom;
        this.left = builder.left;
        this.right = builder.right;
    }

    public static class Builder {

        private BorderStyle top;
        private BorderStyle bottom;
        private BorderStyle left;
        private BorderStyle right;


        public Builder top(BorderStyle topBorderStyle) {
            this.top = topBorderStyle;
            return this;
        }

        public Builder bottom(BorderStyle bottomBorderStyle) {
            this.bottom = bottomBorderStyle;
            return this;
        }

        public Builder left(BorderStyle leftBorderStyle) {
            this.left = leftBorderStyle;
            return this;
        }

        public Builder right(BorderStyle rightBorderStyle) {
            this.right = rightBorderStyle;
            return this;
        }

        public DefaultExcelBorder build() {
            return new DefaultExcelBorder(this);
        }
    }

    @Override
    public void applyAllBorder(CellStyle cellStyle) {
        if (top != null) cellStyle.setBorderTop(top.getBorderStyle());
        if (bottom != null) cellStyle.setBorderBottom(bottom.getBorderStyle());
        if (left != null) cellStyle.setBorderLeft(left.getBorderStyle());
        if (right != null) cellStyle.setBorderRight(right.getBorderStyle());
    }
}
