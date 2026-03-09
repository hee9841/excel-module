package io.github.hee9841.excel.core.exporter;

import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.mode.SheetMode;
import java.text.MessageFormat;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * SXSSFExporter is a concrete implementation of {@link AbstractExcelExporter} that provides functionality
 * for exporting data to Excel files. This class uses the SXSSFWorkbook from Apache POI for
 * efficient handling of large datasets by streaming data to disk.
 *
 * <p>The SXSSFExporter supports two sheet management modes:</p>
 * <ul>
 *     <li>ONE_SHEET - All data is exported to a single sheet (limited by max rows per sheet)</li>
 *     <li>MULTI_SHEET - Data is split across multiple sheets when exceeding max rows per sheet</li>
 * </ul>
 *
 * <p>Use the {@link SXSSFExporterBuilder} to configure and instantiate this class.</p>
 *
 * @param <T> The type of data to be exported to Excel. The type must be annotated appropriately
 *            for Excel column mapping using the library's annotation system.
 * @see ExcelExporter
 * @see AbstractExcelExporter
 * @see SXSSFExporterBuilder
 * @see SheetMode
 */
public class SXSSFExporter<T> extends AbstractExcelExporter<T, SXSSFWorkbook> {

    private static final SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;
    private static final int HEADER_ROW_INDEX = 0;

    private final String sheetNamePrefix;
    private final int maxRowsIndexPerSheet;

    private SheetMode sheetMode;

    private Sheet currentSheet;
    private int currentRowIndex = HEADER_ROW_INDEX;
    private int currentSheetIndex;


    /**
     * Constructs an SXSSFExporter with the specified configuration.
     *
     * <p>This constructor is not meant to be called directly. Use {@link SXSSFExporterBuilder}
     * to create instances of SXSSFExporter.</p>
     *
     * @param type            The class type of the data to be exported
     * @param data            The list of data objects to be exported
     * @param sheetMode       The mode for sheet management (ONE_SHEET or MULTI_SHEET)
     * @param sheetName       Base name for sheets (null for default names)
     * @param maxRowsPerSheet Maximum number of rows allowed per sheet
     */
    SXSSFExporter(
        Class<T> type,
        List<T> data,
        SheetMode sheetMode,
        String sheetName,
        int maxRowsPerSheet
    ) {
        super(new SXSSFWorkbook());

        this.sheetNamePrefix = sheetName;
        this.maxRowsIndexPerSheet = maxRowsPerSheet - 1;
        this.currentSheetIndex = HEADER_ROW_INDEX;

        setSheetMode(sheetMode);
        // Initialize column mapping
        this.initialize(type, data);
        //create Excel
        this.buildExcel(data);
    }


    /**
     * Creates a new builder for configuring and instantiating an SXSSFExporter.
     *
     * @param <T>  The type of data to be exported
     * @param type The class of the data type
     * @param data The list of data objects to be exported
     * @return A new SXSSFExporterBuilder instance
     */
    public static <T> SXSSFExporterBuilder<T> builder(Class<T> type, List<T> data) {
        return new SXSSFExporterBuilder<>(type, data, supplyExcelVersion.getMaxRows());
    }

    /**
     * Sets the sheet mode for this exporter.
     *
     * <p>This method also configures the workbook's Zip64 mode based on the selected mode.</p>
     *
     * @param mode The sheet mode to use (ONE_SHEET or MULTI_SHEET)
     */
    private void setSheetMode(SheetMode mode) {

        this.sheetMode = mode;
        workbook.setZip64Mode(sheetMode.getZip64Mode());

        logger.debug("Set sheet mode and Zip64Mode - mode: {}, Zip64Mode: {}.",
            mode.name(), sheetMode.getZip64Mode().name());
    }


    /**
     * Validates the data size against the maximum rows per sheet limit.
     *
     * <p>This method checks if the data size exceeds the maximum allowed rows per sheet
     * when using ONE_SHEET mode. If the limit is exceeded, an ExcelException is thrown.</p>
     *
     * @param type The class type of the data being validated
     * @param data The list of data objects to be validated
     * @throws ExcelException if data size exceeds max rows limit with ONE_SHEET mode
     */
    @Override
    protected void validate(Class<?> type, List<T> data) {
        if (type == null) {
            throw new IllegalArgumentException("Type must not be null.");
        }
        if (data == null) {
            throw new IllegalArgumentException("Data must not be null.");
        }

        if (SheetMode.isOneSheet(sheetMode) && data.size() > maxRowsIndexPerSheet) {
            throw new ExcelException(
                MessageFormat.format(
                    "The data size exceeds the maximum number of data rows allowed per sheet. "
                        + "The sheet mode is set to ONE_SHEET but the data size is larger than "
                        + "the maximum data rows per sheet (excluding header) (data size: {0}, maximum data rows: {1}).\n"
                        + "Please change the sheet mode to MULTI_SHEET or reduce the data size.",
                    data.size(), maxRowsIndexPerSheet
                ), dtoTypeName);
        }
    }

    /**
     * Adds rows to the current sheet for the provided data list.
     *
     * <p>If the number of rows exceeds the maximum allowed per sheet and the sheet mode
     * is MULTI_SHEET, a new sheet will be created to continue adding rows.</p>
     *
     * <p>If the sheet mode is ONE_SHEET and the data size exceeds the remaining rows
     * in the current sheet, an ExcelException will be thrown.</p>
     *
     * @param data The list of data objects to be added as rows
     * @throws ExcelException if ONE_SHEET mode is used and data exceeds max rows limit
     */
    @Override
    protected void doAddRows(List<T> data) {
        // If sheet mode ONE_SHEET and ata size exceeds the remaining rows, throw Exception
        if (SheetMode.isOneSheet(sheetMode) &&
            (data.size() > maxRowsIndexPerSheet - currentRowIndex)
        ) {
            throw new ExcelException(
                MessageFormat.format(
                    "The data size exceeds the remaining data rows in the current sheet. "
                        + "The sheet mode is set to ONE_SHEET but the data size is larger than "
                        + "the remaining data rows (data size: {0}, remaining data rows: {1}, maximum data rows per sheet (excluding header): {2}).\n"
                        + "Please change the sheet mode to MULTI_SHEET or reduce the data size.",
                    data.size(), (maxRowsIndexPerSheet - currentRowIndex), maxRowsIndexPerSheet),
                dtoTypeName);
        }


        for (T rowData : data) {
            if (currentRowIndex >= maxRowsIndexPerSheet) {
                initializeNewSheet();
            }
            createRow(currentSheet, rowData, ++currentRowIndex);
        }
    }


    private void initializeNewSheet() {
        currentSheet = createSheet(sheetNamePrefix, currentSheetIndex++);
        currentRowIndex = HEADER_ROW_INDEX;
        createHeader(currentSheet, currentRowIndex);
    }

    /**
     * Creates the first Excel(workBook) with the provided data.
     *
     * <p>This method handles the creation of sheets and rows based on the data:</p>
     * <ul>
     *   <li>If the data is empty, it creates a sheet with headers only</li>
     *   <li>Otherwise, it creates a sheet with headers and adds all data rows</li>
     * </ul>
     *
     * @param data The list of data objects to be exported
     */
    private void buildExcel(List<T> data) {
        try {
            // Initialize first sheet with headers
            initializeNewSheet();

            // 1. If data is empty, create createHeader only.
            if (data.isEmpty()) {
                logger.warn("Empty data provided - Excel file will be created with headers only.");
                return;
            }

            //2. Add Rows
            doAddRows(data);
        } catch (Exception e) {
            close();
            throw e;
        }
    }





}
