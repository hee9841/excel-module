package io.github.hee9841.excel.meta;

import org.apache.poi.ss.usermodel.CellStyle;

public class ColumnInfo {

    private final String fieldName;
    private final String headerName;
    private final CellType columnType;
    private final CellStyle headerStyle;
    private final CellStyle bodyStyle;

    private ColumnInfo(
        String fieldName,
        String headerName,
        CellType columnType,
        CellStyle headerStyle,
        CellStyle bodyStyle
    ) {
        this.fieldName = fieldName;
        this.headerName = headerName;
        this.columnType = columnType;
        this.headerStyle = headerStyle;
        this.bodyStyle = bodyStyle;
    }

    public static ColumnInfo of(
        String fieldName,
        String headerName,
        CellType columnType,
        CellStyle headerStyle,
        CellStyle bodyStyle
    ) {
        return new ColumnInfo(fieldName, headerName, columnType, headerStyle, bodyStyle);
    }

    public String getFieldName() {
        return fieldName;
    }

    public String getHeaderName() {
        return headerName;
    }

    public CellType getColumnType() {
        return columnType;
    }

    public CellStyle getHeaderStyle() {
        return headerStyle;
    }

    public CellStyle getBodyStyle() {
        return bodyStyle;
    }
}
