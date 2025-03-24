package io.github.hee9841.excel.style.border;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Implementation of ExcelBorder that allows for configuring individual border styles
 * for each side of an Excel cell (top, bottom, left, right).
 * 
 * <pre>
 * This class provides multiple ways to create border configurations:
 * - Static factory method for applying the same border style to all sides
 * - Builder pattern for configuring each side individually
 * </pre>
 */
public class DefaultExcelBorder implements ExcelBorder {

    private final BorderStyle top;
    private final BorderStyle bottom;
    private final BorderStyle left;
    private final BorderStyle right;

    /**
     * Creates a DefaultExcelBorder with the same border style applied to all sides.
     *
     * @param all the border style to apply to all sides
     * @return a new DefaultExcelBorder instance
     */
    public static DefaultExcelBorder all(BorderStyle all) {
        return new DefaultExcelBorder(all, all, all, all);
    }

    /**
     * Creates a new Builder instance to construct a DefaultExcelBorder.
     *
     * @return a new Builder instance
     */
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

    /**
     * Builder class for creating DefaultExcelBorder instances with custom
     * border styles for each side.
     */
    public static class Builder {

        private BorderStyle top;
        private BorderStyle bottom;
        private BorderStyle left;
        private BorderStyle right;

        /**
         * Sets the border style for the top side.
         *
         * @param topBorderStyle the border style for the top side
         * @return this builder instance
         */
        public Builder top(BorderStyle topBorderStyle) {
            this.top = topBorderStyle;
            return this;
        }

        /**
         * Sets the border style for the bottom side.
         *
         * @param bottomBorderStyle the border style for the bottom side
         * @return this builder instance
         */
        public Builder bottom(BorderStyle bottomBorderStyle) {
            this.bottom = bottomBorderStyle;
            return this;
        }

        /**
         * Sets the border style for the left side.
         *
         * @param leftBorderStyle the border style for the left side
         * @return this builder instance
         */
        public Builder left(BorderStyle leftBorderStyle) {
            this.left = leftBorderStyle;
            return this;
        }

        /**
         * Sets the border style for the right side.
         *
         * @param rightBorderStyle the border style for the right side
         * @return this builder instance
         */
        public Builder right(BorderStyle rightBorderStyle) {
            this.right = rightBorderStyle;
            return this;
        }

        /**
         * Builds a new DefaultExcelBorder instance.
         *
         * @return a new DefaultExcelBorder instance
         */
        public DefaultExcelBorder build() {
            return new DefaultExcelBorder(this);
        }
    }

    /**
     * Applies the configured border styles to the given cell style.
     * If any border style is null, that particular side will not have a border applied.
     *
     * @param cellStyle the cell style to which borders will be applied
     */
    @Override
    public void applyAllBorder(CellStyle cellStyle) {
        if (top != null) cellStyle.setBorderTop(top.getBorderStyle());
        if (bottom != null) cellStyle.setBorderBottom(bottom.getBorderStyle());
        if (left != null) cellStyle.setBorderLeft(left.getBorderStyle());
        if (right != null) cellStyle.setBorderRight(right.getBorderStyle());
    }
}
