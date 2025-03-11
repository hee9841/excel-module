package io.github.hee9841.excel.style.border;

public enum BorderStyle {
    NONE(org.apache.poi.ss.usermodel.BorderStyle.NONE),
    THIN(org.apache.poi.ss.usermodel.BorderStyle.THIN),
    MEDIUM(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM),
    DASHED(org.apache.poi.ss.usermodel.BorderStyle.DASHED),
    DOTTED(org.apache.poi.ss.usermodel.BorderStyle.DOTTED),
    THICK(org.apache.poi.ss.usermodel.BorderStyle.THICK),
    DOUBLE(org.apache.poi.ss.usermodel.BorderStyle.DOUBLE),
    HAIR(org.apache.poi.ss.usermodel.BorderStyle.HAIR),
    MEDIUM_DASHED(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASHED),
    DASH_DOT(org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT),
    MEDIUM_DASH_DOT(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT),
    DASH_DOT_DOT(org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT_DOT),
    MEDIUM_DASH_DOT_DOT(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT_DOT),
    SLANTED_DASH_DOT(org.apache.poi.ss.usermodel.BorderStyle.SLANTED_DASH_DOT);

    private final org.apache.poi.ss.usermodel.BorderStyle borderStyle;

    BorderStyle(org.apache.poi.ss.usermodel.BorderStyle borderStyle) {
        this.borderStyle = borderStyle;
    }

    public org.apache.poi.ss.usermodel.BorderStyle getBorderStyle() {
        return this.borderStyle;
    }
}
