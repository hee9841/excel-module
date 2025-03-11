package io.github.hee9841.excel.meta;

public class ColumnInfo {

    private final String fieldName;
    private final String headerName;

    private ColumnInfo(
        String fieldName,
        String headerName
    ) {
        this.fieldName = fieldName;
        this.headerName = headerName;
    }

    public static ColumnInfo of(
        String fieldName,
        String headerName
    ) {
        return new ColumnInfo(fieldName, headerName);
    }

    public String getFieldName() {
        return fieldName;
    }

    public String getHeaderName() {
        return headerName;
    }
}
