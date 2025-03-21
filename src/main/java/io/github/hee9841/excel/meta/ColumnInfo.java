package io.github.hee9841.excel.meta;

import io.github.hee9841.excel.annotation.ExcelColumnStyle;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Represents information about an Excel column, containing field mapping and style details.
 * This class holds the necessary information to map a Java field to an Excel column, 
 * including styling for header and body cells.
 * 
 * @see ColumnInfoMapper
 * @see CellType
 * @see io.github.hee9841.excel.annotation.Excel
 * @see io.github.hee9841.excel.annotation.ExcelColumn
 * @see ExcelColumnStyle
 */
public class ColumnInfo {

    /** The name of the Java field this column maps to */
    private final String fieldName;
    /** The display name to use in the Excel header */
    private final String headerName;
    /** The cell type to use for this column */
    private final CellType columnType;
    /** The cell style to apply to the header cell */
    private final CellStyle headerStyle;
    /** The cell style to apply to body cells in this column */
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

    /**
     * Factory method to create a new {@link ColumnInfo} instance.
     *
     * @param fieldName The name of the Java field
     * @param headerName The display name for the Excel header
     * @param columnType The {@link CellType} for this column
     * @param headerStyle The style for the header cell
     * @param bodyStyle The style for the body cells
     * @return A new {@link ColumnInfo} instance
     */
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
