package io.github.hee9841.excel.core;

import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.strategy.SheetStrategy;
import java.text.MessageFormat;
import java.util.List;

/**
 * ExcelExporter is a concrete implementation of {@link SXSSFExcelFile} that provides functionality
 * for exporting data to Excel files. This class uses the SXSSFWorkbook from Apache POI for efficient
 * handling of large datasets by streaming data to disk.
 * 
 * <p>The ExcelExporter supports two sheet management strategies:</p>
 * <ul>
 *     <li>ONE_SHEET - All data is exported to a single sheet (limited by max rows per sheet)</li>
 *     <li>MULTI_SHEET - Data is split across multiple sheets when exceeding max rows per sheet</li>
 * </ul>
 * 
 * <p>Use the {@link ExcelExporterBuilder} to configure and instantiate this class.</p>
 * 
 * @param <T> The type of data to be exported to Excel. The type must be annotated appropriately
 *           for Excel column mapping using the library's annotation system.
 * 
 * @see SXSSFExcelFile
 * @see ExcelExporterBuilder
 * @see SheetStrategy
 */
public class ExcelExporter<T> extends SXSSFExcelFile<T> {

    private static final String EXCEED_MAX_ROW_MSG_2ARGS =
        "The data size exceeds the maximum number of rows allowed per sheet. "
            + "The sheet strategy is set to ONE_SHEET but the data size ({0} rows) is larger than "
            + "the maximum rows per sheet ({1} rows).\n"
            + "Please change the sheet strategy to MULTI_SHEET or reduce the data size.";

    private static final int ROW_START_INDEX = 0;
    private int currentRowIndex = ROW_START_INDEX;

    private final String sheetName;
    private final int maxRowsPerSheet;

    private SheetStrategy sheetStrategy;
    private int currentSheetIndex;


    /**
     * Constructs an ExcelExporter with the specified configuration.
     * 
     * <p>This constructor is not meant to be called directly. Use {@link ExcelExporterBuilder}
     * to create instances of ExcelExporter.</p>
     *
     * @param type           The class type of the data to be exported
     * @param data           The list of data objects to be exported
     * @param sheetStrategy  The strategy for sheet management (ONE_SHEET or MULTI_SHEET)
     * @param sheetName      Base name for sheets (null for default names)
     * @param maxRowsPerSheet Maximum number of rows allowed per sheet
     */
    ExcelExporter(
        Class<T> type,
        List<T> data,
        SheetStrategy sheetStrategy,
        String sheetName,
        int maxRowsPerSheet
    ) {
        super();
        this.maxRowsPerSheet = maxRowsPerSheet;
        this.sheetName = sheetName;
        this.currentSheetIndex = 0;

        this.initialize(type, data);
        setSheetStrategy(sheetStrategy, data.size());
        this.createExcelFile(data);
    }


    /**
     * Creates a new builder for configuring and instantiating an ExcelExporter.
     *
     * @param <T>  The type of data to be exported
     * @param type The class of the data type
     * @param data The list of data objects to be exported
     * @return A new ExcelExporterBuilder instance
     */
    public static <T> ExcelExporterBuilder<T> builder(Class<T> type, List<T> data) {
        return new ExcelExporterBuilder<>(type, data, supplyExcelVersion.getMaxRows());
    }


    /**
     * Creates the Excel file with the provided data.
     * 
     * <p>This method handles the creation of sheets and rows based on the data:</p>
     * <ul>
     *   <li>If the data is empty, it creates a sheet with headers only</li>
     *   <li>Otherwise, it creates a sheet with headers and adds all data rows</li>
     * </ul>
     *
     * @param data The list of data objects to be exported
     */
    @Override
    protected void createExcelFile(List<T> data) {
        // 1. If data is empty, create createHeader only.
        if (data.isEmpty()) {
            createNewSheetWithHeader();
            logger.warn("Empty data provided - Excel file will be created with headers only.");
            return;
        }

        //2. Add rows
        createNewSheetWithHeader();
        addRows(data);
    }

    /**
     * Adds rows to the current sheet for the provided data list.
     * 
     * <p>If the number of rows exceeds the maximum allowed per sheet and the sheet strategy
     * is MULTI_SHEET, a new sheet will be created to continue adding rows.</p>
     * 
     * <p>If the sheet strategy is ONE_SHEET and the data size exceeds the maximum rows per sheet,
     * an ExcelException will be thrown.</p>
     *
     * @param data The list of data objects to be added as rows
     * @throws ExcelException if ONE_SHEET strategy is used and data exceeds max rows limit
     */
    @Override
    public void addRows(List<T> data) {
        int leftDataSize = data.size();
        for (Object renderedData : data) {
            createBody(renderedData, currentRowIndex++);
            leftDataSize--;
            if (currentRowIndex == maxRowsPerSheet && leftDataSize > 0) {
                //If one sheet strategy, throw exception
                if (SheetStrategy.isOneSheet(sheetStrategy)) {
                    throw new ExcelException(
                        MessageFormat.format(EXCEED_MAX_ROW_MSG_2ARGS,
                            currentRowIndex + data.size(), maxRowsPerSheet), dtoTypeName);
                }

                //If multi sheet strategy, create new sheet
                createNewSheetWithHeader();
            }
        }
    }

    /**
     * Sets the sheet strategy for this exporter and validates that it is compatible with the data size.
     * 
     * <p>This method also configures the workbook's Zip64 mode based on the selected strategy.</p>
     *
     * @param strategy The sheet strategy to use (ONE_SHEET or MULTI_SHEET)
     * @param dataSize The size of the data to be exported
     * @throws ExcelException if ONE_SHEET strategy is used and data size exceeds max rows limit
     */
    private void setSheetStrategy(SheetStrategy strategy, int dataSize) {
        
        if (dataSize > maxRowsPerSheet && SheetStrategy.isOneSheet(strategy)) {
            throw new ExcelException(
                MessageFormat.format(EXCEED_MAX_ROW_MSG_2ARGS, dataSize, maxRowsPerSheet),
                dtoTypeName);
        }

        this.sheetStrategy = strategy;
        workbook.setZip64Mode(sheetStrategy.getZip64Mode());

        logger.debug("Set sheet strategy and Zip64Mode - strategy: {}, Zip64Mode:{}.",
            strategy.name(), sheetStrategy.getZip64Mode().name());
    }

    /**
     * Creates a new sheet with headers.
     * 
     * <p>This method resets the current row index, creates a new sheet, and adds headers to it.
     * If a sheet name is provided, it will be used as a base name with an index(index starts from 0) suffix.</p>
     */
    private void createNewSheetWithHeader() {
        currentRowIndex = ROW_START_INDEX;

        //If sheet name is provided, create sheet with sheet name + idx
        if (sheetName != null) {
            sheet = workbook.createSheet(String.format("%s%d", sheetName, currentSheetIndex++));
        } else {
            sheet = workbook.createSheet();
        }

        logger.debug("Create new Sheet : {}.", sheet.getSheetName());

        createHeaderWithNewSheet(sheet, ROW_START_INDEX);
        currentRowIndex++;
    }
}
