package io.github.hee9841.excel.core.exporter;

import io.github.hee9841.excel.core.meta.ColumnInfo;
import io.github.hee9841.excel.core.meta.ColumnInfoMapper;
import io.github.hee9841.excel.exception.ExcelException;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Abstract base class for Excel file operations using Apache POI's Workbook abstraction.
 * This class provides the core functionality for handling Excel files with different Workbook
 * implementations.
 *
 * <p>Key features:</p>
 * <ul>
 *     <li>Supports different Workbook implementations</li>
 *     <li>Supports Excel 2007+ format (XLSX) by default</li>
 *     <li>Provides column mapping and header generation</li>
 *     <li>Handles cell styling and data type conversion</li>
 * </ul>
 *
 * <p>This class implements the core functionality while leaving sheet management
 * to be implemented by concrete subclasses based on their configured {@link io.github.hee9841.excel.mode.SheetMode}.</p>
 *
 * <p><b>Thread Safety:</b> This class is NOT thread-safe. A single instance should not be
 * shared across multiple threads. Each thread should create its own exporter instance.</p>
 *
 * @param <T> The type of data to be handled in the Excel file
 * @param <W> The workbook implementation type
 */
public abstract class AbstractExcelExporter<T, W extends Workbook> implements ExcelExporter<T> {

    // ========== Fields ==========

    protected static final Logger logger = LoggerFactory.getLogger(AbstractExcelExporter.class);

    protected W workbook;
    protected List<ColumnInfo> columnsMappingInfos;
    protected String dtoTypeName;

    private volatile boolean closed = false;

    // ========== Constructor ==========

    /**
     * Constructs a new AbstractExcelExporter with the provided workbook instance.
     */
    protected AbstractExcelExporter(W workbook) {
        if (workbook == null) {
            throw new ExcelException("Workbook cannot be null");
        }
        this.workbook = workbook;
    }

    // ========== Public API ==========

    /**
     * Adds additional rows to the existing Excel file.
     * This method checks if the exporter is still open before delegating to
     * the subclass implementation.
     *
     * @param data The list of data objects to be added as rows
     * @throws IllegalStateException if the exporter has already been closed
     */
    @Override
    public final void addRows(List<T> data) {
        ensureOpen();
        doAddRows(data);
    }

    /**
     * Writes the Excel file content to the specified output stream.
     * After this method is called, the exporter is closed and cannot be reused.
     *
     * @param stream The output stream to write the Excel file to
     * @throws IOException if an I/O error occurs during writing
     * @throws IllegalArgumentException if stream is null
     * @throws IllegalStateException if the exporter has already been closed
     */
    @Override
    public final void write(OutputStream stream) throws IOException {
        ensureOpen();

        if (stream == null) {
            throw new IllegalArgumentException("Output stream must not be null.");
        }
        logger.info("Start to write Excel file for DTO class({}.java).", dtoTypeName);

        try {
            workbook.write(stream);
            logger.info("Successfully wrote Excel file for DTO class({}.java).", dtoTypeName);
        } finally {
            close();
        }
    }


    /**
     * Closes the exporter and releases any resources associated with it.
     * This method is idempotent - calling it multiple times has no additional effect.
     *
     * <p>For {@code SXSSFWorkbook}, this also cleans up temporary files created during
     * the streaming process.</p>
     */
    @Override
    public final void close() {
        if (closed) {
            return;
        }
        try {
            workbook.close();
            logger.debug("Workbook closed for DTO class({}.java).", dtoTypeName);
        } catch (IOException e) {
            logger.warn("Failed to close workbook for DTO class({}.java).", dtoTypeName, e);
        } finally {
            closed = true;
        }
    }

    // ========== Template Methods ==========

    /**
     * Initializes the Excel file with the specified type and data.
     * This method performs validation and sets up column mapping information.
     *
     * @param type The class type of the data to be exported
     * @param data The list of data objects to be exported
     */
    protected final void initialize(Class<?> type, List<T> data) {
        if (type == null) {
            throw new IllegalArgumentException("Type must not be null.");
        }
        if (data == null) {
            throw new IllegalArgumentException("Data must not be null.");
        }

        this.dtoTypeName = type.getName();
        logger.info("Initializing Excel file for DTO: {}.java.", dtoTypeName);

        validate(type, data);

        logger.debug("Mapping DTO to Excel data - DTO class({}).", dtoTypeName);
        //Map DTO to Excel data
        this.columnsMappingInfos = ColumnInfoMapper.of(type, workbook).map();
    }

    // ========== Abstract Methods (to be implemented by subclasses) ==========

    /**
     * Validates the provided data and type.
     * This method can be overridden by subclasses to add custom validation logic.
     *
     * @param type The class of the data type
     * @param data The list of data objects to be exported
     */
    protected abstract void validate(Class<?> type, List<T> data);

    /**
     * Creates the Excel file with the provided data.
     * This method must be implemented by subclasses to define their specific sheet management
     * based on the configured {@link io.github.hee9841.excel.mode.SheetMode}.
     *
     * @param data The list of data objects to be exported
     */
    protected abstract void createExcel(List<T> data);

    /**
     * Performs the actual row addition logic.
     * This method must be implemented by subclasses according to their configured
     * {@link io.github.hee9841.excel.mode.SheetMode} and workbook type.
     *
     * @param data The list of data objects to be added as rows
     */
    protected abstract void doAddRows(List<T> data);

    // ========== Helper Methods ==========

    /**
     * Creates a new sheet.
     *
     * <p> If a sheet name is provided, it will be used as a base name with an index (index starts
     * from 0) suffix.</p>
     *
     * @param sheetNamePrefix Base name for sheets (null for default names)
     * @param sheetIndex      Index used to suffix the sheet name
     * @return The newly created sheet
     */
    protected final Sheet createSheet(String sheetNamePrefix, int sheetIndex) {
        //If sheet name is provided, create sheet with sheet name + idx
        final String finalSheetName = (sheetNamePrefix != null)
            ? String.format("%s(%d)", sheetNamePrefix, sheetIndex)
            : null;


        Sheet sheet = (finalSheetName != null)
            ? workbook.createSheet(finalSheetName)
            : workbook.createSheet();

        logger.debug("Create new Sheet : {}.", sheet.getSheetName());

        return sheet;
    }


    /**
     * Creates a header row using the column mapping information.
     *
     * @param sheet The sheet to add headers to
     * @param headerRowIndex The row index where headers should be created (0-based)
     */
    protected final void createHeader(Sheet sheet, Integer headerRowIndex) {
        Row row = sheet.createRow(headerRowIndex);
        for (ColumnInfo columnMappingInfo : columnsMappingInfos) {
            int colIndex = columnMappingInfo.getIndex();
            Cell cell = row.createCell(colIndex);
            cell.setCellValue(columnMappingInfo.getHeaderName());
            cell.setCellStyle(columnMappingInfo.getHeaderStyle());
        }
        logger.debug("Created header row at index {}", headerRowIndex);
    }

    /**
     * Creates a row in the Excel sheet for the given data object.
     * This method handles field access and cell value setting based on column mapping information.
     *
     * @param sheet    The Sheet object to create a row.
     * @param data     The data object for rendering data to cell
     * @param rowIndex The index of the row to create
     * @throws ExcelException if field access fails
     */
    protected final void createRow(Sheet sheet, Object data, int rowIndex) {
        logger.debug("Add rows data - row:{}.", rowIndex);
        Row row = sheet.createRow(rowIndex);

        for (ColumnInfo columnInfo : columnsMappingInfos) {
            int colIndex = columnInfo.getIndex();
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

    // ========== Private Methods ==========

    /**
     * Ensures that the exporter is still open and can be used.
     *
     * @throws IllegalStateException if the exporter has already been closed
     */
    private void ensureOpen() {
        if (closed) {
            throw new IllegalStateException(
                "Exporter is already closed. Cannot reuse after write().");
        }
    }

}
