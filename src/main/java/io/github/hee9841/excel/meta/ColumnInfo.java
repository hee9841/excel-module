package io.github.hee9841.excel.meta;

public class ColumnInfo {

    private final String fieldName;
    private final String headerName;
    private final CellType columnType;

    private ColumnInfo(
        String fieldName,
        String headerName,
        CellType columnType
    ) {
        this.fieldName = fieldName;
        this.headerName = headerName;
        this.columnType = columnType;
    }

    public static ColumnInfo of(
        String fieldName,
        String headerName,
        CellType columnType
    ) {
        return new ColumnInfo(fieldName, headerName, columnType);
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
}
