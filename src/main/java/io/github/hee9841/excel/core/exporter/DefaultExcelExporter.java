package io.github.hee9841.excel.core.exporter;

import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.strategy.SheetStrategy;
import java.text.MessageFormat;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * DefaultExcelExporter is a concrete implementation of {@link SXSSFExporter} that provides functionality
 * for exporting data to Excel files. This class uses the SXSSFWorkbook from Apache POI for
 * efficient
 * handling of large datasets by streaming data to disk.
 *
 * <p>The DefaultExcelExporter supports two sheet management strategies:</p>
 * <ul>
 *     <li>ONE_SHEET - All data is exported to a single sheet (limited by max rows per sheet)</li>
 *     <li>MULTI_SHEET - Data is split across multiple sheets when exceeding max rows per sheet</li>
 * </ul>
 *
 * <p>Use the {@link DefaultExcelExporterBuilder} to configure and instantiate this class.</p>
 *
 * @param <T> The type of data to be exported to Excel. The type must be annotated appropriately
 *            for Excel column mapping using the library's annotation system.
 * @see SXSSFExporter
 * @see DefaultExcelExporterBuilder
 * @see SheetStrategy
 */
public class DefaultExcelExporter<T> extends SXSSFExporter<T> {

    private static final String EXCEED_MAX_ROW_MSG_2ARGS =
        "The data size exceeds the maximum number of rows allowed per sheet. "
            + "The sheet strategy is set to ONE_SHEET but the data size is larger than "
            + "the maximum rows per sheet (data size: {0}, maximum rows: {1} ).\n"
            + "Please change the sheet strategy to MULTI_SHEET or reduce the data size.";

    private static final int ROW_START_INDEX = 0;
    private int currentRowIndex = ROW_START_INDEX;

    private final String sheetName;
    private final int maxRowsPerSheet;

    private SheetStrategy sheetStrategy;

    private Sheet currentSheet;


    /**
     * Constructs an DefaultExcelExporter with the specified configuration.
     *
     * <p>This constructor is not meant to be called directly. Use {@link DefaultExcelExporterBuilder}
     * to create instances of DefaultExcelExporter.</p>
     *
     * @param type            The class type of the data to be exported
     * @param data            The list of data objects to be exported
     * @param sheetStrategy   The strategy for sheet management (ONE_SHEET or MULTI_SHEET)
     * @param sheetName       Base name for sheets (null for default names)
     * @param maxRowsPerSheet Maximum number of rows allowed per sheet
     */
    DefaultExcelExporter(
        Class<T> type,
        List<T> data,
        SheetStrategy sheetStrategy,
        String sheetName,
        int maxRowsPerSheet
    ) {
        super();
        this.maxRowsPerSheet = maxRowsPerSheet;
        this.sheetName = sheetName;
        setSheetStrategy(sheetStrategy);

        this.initialize(type, data);
        this.createExcel(data);
    }


    /**
     * Creates a new builder for configuring and instantiating an DefaultExcelExporter.
     *
     * @param <T>  The type of data to be exported
     * @param type The class of the data type
     * @param data The list of data objects to be exported
     * @return A new DefaultExcelExporterBuilder instance
     */
    public static <T> DefaultExcelExporterBuilder<T> builder(Class<T> type, List<T> data) {
        return new DefaultExcelExporterBuilder<>(type, data, supplyExcelVersion.getMaxRows());
    }

    /**
     * Sets the sheet strategy for this exporter.
     *
     * <p>This method also configures the workbook's Zip64 mode based on the selected strategy.</p>
     *
     * @param strategy The sheet strategy to use (ONE_SHEET or MULTI_SHEET)-
     */
    private void setSheetStrategy(SheetStrategy strategy) {

        this.sheetStrategy = strategy;
        workbook.setZip64Mode(sheetStrategy.getZip64Mode());

        logger.debug("Set sheet strategy and Zip64Mode - strategy: {}, Zip64Mode: {}.",
            strategy.name(), sheetStrategy.getZip64Mode().name());
    }


    /**
     * Validates the data size against the maximum rows per sheet limit.
     *
     * <p>This method checks if the data size exceeds the maximum allowed rows per sheet
     * when using ONE_SHEET strategy. If the limit is exceeded, an ExcelException is thrown.</p>
     *
     * @param type The class type of the data being validated
     * @param data The list of data objects to be validated
     * @throws ExcelException if data size exceeds max rows limit with ONE_SHEET strategy
     */
    @Override
    protected void validate(Class<?> type, List<T> data) {
        if (SheetStrategy.isOneSheet(sheetStrategy) && data.size() > maxRowsPerSheet - 1) {
            throw new ExcelException(
                MessageFormat.format(EXCEED_MAX_ROW_MSG_2ARGS,
                    data.size(), maxRowsPerSheet
                ), dtoTypeName);
        }
    }

    /**
     * Creates the Excel(workBook) with the provided data.
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
    protected void createExcel(List<T> data) {

        currentSheet = createNewSheet(sheetName, 0);
        createHeader(currentSheet, ROW_START_INDEX);

        // 1. If data is empty, create createHeader only.
        if (data.isEmpty()) {
            logger.warn("Empty data provided - Excel file will be created with headers only.");
            return;
        }

        //2. Add Rows
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
            createBody(currentSheet, renderedData, currentRowIndex++);
            leftDataSize--;
            if (currentRowIndex == maxRowsPerSheet && leftDataSize > 0) {
                //If one sheet strategy, throw exception
                if (SheetStrategy.isOneSheet(sheetStrategy)) {
                    throw new ExcelException(
                        MessageFormat.format(EXCEED_MAX_ROW_MSG_2ARGS,
                            data.size(), maxRowsPerSheet), dtoTypeName);
                }

                //If multi sheet strategy, create new sheet
                currentRowIndex = ROW_START_INDEX;
                currentSheet = createNewSheet(sheetName, workbook.getSheetIndex(currentSheet) + 1);
                createHeader(currentSheet, ROW_START_INDEX);
            }
        }
    }

    /**
     * Override createHeader Method to add currentRowIndex.
     *
     * @param sheet      The sheet to add headers to
     * @param headerRowIndex The headers row index
     */
    @Override
    protected void createHeader(Sheet sheet, Integer headerRowIndex) {
        super.createHeader(sheet, headerRowIndex);
        currentRowIndex++;
    }
}
