package io.github.hee9841.excel.core;

import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.meta.ColumnInfo;
import io.github.hee9841.excel.meta.ColumnInfoMapper;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Abstract base class for Excel file operations using Apache POI's SXSSF (Streaming XML Spreadsheet Format).
 * This class provides the core functionality for handling Excel files with streaming support for large datasets.
 *
 * <p>Key features:</p>
 * <ul>
 *     <li>Uses SXSSFWorkbook for memory-efficient handling of large Excel files</li>
 *     <li>Supports Excel 2007+ format (XLSX)</li>
 *     <li>Provides column mapping and header generation</li>
 *     <li>Handles cell styling and data type conversion</li>
 * </ul>
 *
 * <p>This class implements the core functionality while leaving sheet management strategies
 * to be implemented by concrete subclasses.</p>
 *
 * @param <T> The type of data to be handled in the Excel file
 */
public abstract class SXSSFExcelFile<T> implements ExcelFile<T> {

    protected static final Logger logger = LoggerFactory.getLogger(SXSSFExcelFile.class);

    protected static final SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;

    protected SXSSFWorkbook workbook;
    protected Map<Integer, ColumnInfo> columnsMappingInfo;
    protected Sheet sheet;

    protected String dtoTypeName;

    /**
     * Constructs a new SXSSFExcelFile with a new SXSSFWorkbook instance.
     */
    protected SXSSFExcelFile() {
        this.workbook = new SXSSFWorkbook();
    }

    /**
     * Initializes the Excel file with the specified type and data.
     * This method performs validation and sets up column mapping information.
     *
     * @param type The class type of the data to be exported
     * @param data The list of data objects to be exported
     */
    protected void initialize(Class<?> type, List<T> data) {
        this.dtoTypeName = type.getName();
        logger.info("Initializing Excel file for DTO: {}.java.", dtoTypeName);

        validate(type, data);

        logger.debug("Mapping DTO to Excel data - DTO class({}).", dtoTypeName);
        //Map DTO to Excel data
        this.columnsMappingInfo = ColumnInfoMapper.of(type, workbook).map();
    }

    /**
     * Validates the provided data and type.
     * This method can be overridden by subclasses to add custom validation logic.
     *
     * @param type The class of the data type
     * @param data The list of data objects to be exported
     */
    protected void validate(Class<?> type, List<T> data) {
    }

    /**
     * Creates the Excel file with the provided data.
     * This method must be implemented by subclasses to define their specific sheet management strategy.
     *
     * @param data The list of data objects to be exported
     */
    protected abstract void createExcelFile(List<T> data);

    /**
     * Creates headers for a new sheet using the column mapping information.
     *
     * @param newSheet The sheet to add headers to
     * @param startRowIndex The row index where headers should start
     */
    protected void createHeaderWithNewSheet(Sheet newSheet, Integer startRowIndex) {
        this.sheet = newSheet;
        Row row = sheet.createRow(startRowIndex);
        for (Integer colIndex : columnsMappingInfo.keySet()) {
            ColumnInfo columnMappingInfo = columnsMappingInfo.get(colIndex);
            Cell cell = row.createCell(colIndex);
            cell.setCellValue(columnMappingInfo.getHeaderName());
            cell.setCellStyle(columnMappingInfo.getHeaderStyle());
        }
    }

    /**
     * Creates a row in the Excel sheet for the given data object.
     * This method handles field access and cell value setting based on column mapping information.
     *
     * @param data The data object to create a row for
     * @param rowIndex The index of the row to create
     * @throws ExcelException if field access fails
     */
    protected void createBody(Object data, int rowIndex) {
        logger.debug("Add rows data - row:{}.", rowIndex);
        Row row = sheet.createRow(rowIndex);
        for (Integer colIndex : columnsMappingInfo.keySet()) {
            ColumnInfo columnInfo = columnsMappingInfo.get(colIndex);
            try {
                Field field = FieldUtils.getField(data.getClass(), columnInfo.getFieldName(), true);
                Cell cell = row.createCell(colIndex);
                //Set cell value by cell type
                columnInfo.getColumnType().setCellValueByCellType(cell, field.get(data));
                //Set cell style
                cell.setCellStyle(columnInfo.getBodyStyle());
            } catch (IllegalAccessException e) {
                throw new ExcelException(
                    String.format("Failed to create body(column:%d, row:%d) : "
                            + "Access to field %s failed.",
                        colIndex, rowIndex, columnInfo.getFieldName()), e);
            }
        }
    }

    /**
     * Writes the Excel file content to the specified output stream.
     * This method ensures proper resource cleanup using try-with-resources.
     *
     * @param stream The output stream to write the Excel file to
     * @throws IOException if an I/O error occurs during writing
     */
    @Override
    public void write(OutputStream stream) throws IOException {
        logger.info("Start to write Excel file for DTO class({}.java).", dtoTypeName);
        try (SXSSFWorkbook autoCloseableWb = this.workbook;
            OutputStream autoCloseableStream = stream) {
            autoCloseableWb.write(autoCloseableStream);
        }
        logger.info("Successfully wrote Excel file for DTO class({}.java).", dtoTypeName);
    }
}
