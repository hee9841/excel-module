package io.github.hee9841.excel.style.border;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Enumeration of border styles for Excel cells.
 * Wraps Apache POI's BorderStyle enum to provide API within this library.
 *
 * <pre>
 * The enum names follow the pattern of "{BORDER_STYLE}" to clearly indicate
 * the border style (e.g. NONE means no border).
 * </pre>
 */
public enum ExcelBorderStyle {
    NONE(BorderStyle.NONE),
    THIN(BorderStyle.THIN),
    MEDIUM(BorderStyle.MEDIUM),
    DASHED(BorderStyle.DASHED),
    DOTTED(BorderStyle.DOTTED),
    THICK(BorderStyle.THICK),
    DOUBLE(BorderStyle.DOUBLE),
    HAIR(BorderStyle.HAIR),
    MEDIUM_DASHED(BorderStyle.MEDIUM_DASHED),
    DASH_DOT(BorderStyle.DASH_DOT),
    MEDIUM_DASH_DOT(BorderStyle.MEDIUM_DASH_DOT),
    DASH_DOT_DOT(BorderStyle.DASH_DOT_DOT),
    MEDIUM_DASH_DOT_DOT(BorderStyle.MEDIUM_DASH_DOT_DOT),
    SLANTED_DASH_DOT(BorderStyle.SLANTED_DASH_DOT);

    private final BorderStyle borderStyle;

    ExcelBorderStyle(BorderStyle borderStyle) {
        this.borderStyle = borderStyle;
    }

    /**
     * Returns the underlying Apache POI BorderStyle.
     *
     * @return the Apache POI BorderStyle corresponding to this enum constant
     */
    public BorderStyle getBorderStyle() {
        return this.borderStyle;
    }
}
