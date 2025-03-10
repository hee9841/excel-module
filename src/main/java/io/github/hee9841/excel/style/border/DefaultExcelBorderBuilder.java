package io.github.hee9841.excel.style.border;

public class DefaultExcelBorderBuilder {

    private BorderStyle top;
    private BorderStyle bottom;
    private BorderStyle left;
    private BorderStyle right;

    DefaultExcelBorderBuilder() {
        this.top = null;
        this.bottom = null;
        this.left = null;
        this.right = null;
    }

    public DefaultExcelBorderBuilder top(BorderStyle topBorderStyle) {
        this.top = topBorderStyle;
        return this;
    }

    public DefaultExcelBorderBuilder bottom(BorderStyle bottomBorderStyle) {
        this.bottom = bottomBorderStyle;
        return this;
    }

    public DefaultExcelBorderBuilder left(BorderStyle leftBorderStyle) {
        this.left = leftBorderStyle;
        return this;
    }

    public DefaultExcelBorderBuilder right(BorderStyle rightBorderStyle) {
        this.right = rightBorderStyle;
        return this;
    }

    public DefaultExcelBorder build() {
        return new DefaultExcelBorder(
            this.top,
            this.bottom,
            this.left,
            this.right
        );
    }

}
