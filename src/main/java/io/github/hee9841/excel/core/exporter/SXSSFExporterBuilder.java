package io.github.hee9841.excel.core.exporter;

import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.mode.SheetMode;
import java.util.List;

/**
 * Builder class for creating and configuring {@link SXSSFExporter} instances.
 * This class implements the Builder pattern to provide a fluent interface for
 * configuring Excel export settings.
 *
 * <p>Default configuration:</p>
 * <ul>
 *     <li>Sheet Mode: MULTI_SHEET</li>
 *     <li>Max Rows per Sheet: Excel 2007+ maximum - 1</li>
 *     <li>Sheet Name: null (default sheet names will be used)</li>
 * </ul>
 *
 * <p>Example usage (recommended with try-with-resources):</p>
 * <pre>{@code
 * try (ExcelExporter<MyData> exporter = SXSSFExporter.builder(MyData.class, dataList)
 *         .sheetMode(SheetMode.ONE_SHEET)
 *         .maxRows(10000)
 *         .sheetName("MySheet")
 *         .build()) {
 *     exporter.addRows(moreData);
 *     exporter.write(outputStream);
 * }
 * }</pre>
 *
 * @param <T> The type of data to be exported
 */
public class SXSSFExporterBuilder<T> {

    private final Class<T> type;
    private final List<T> data;

    private final int supplyExcelMaxRows;

    private int maxRowsPerSheet;
    private SheetMode sheetMode;
    private String sheetName;

    /**
     * Constructs a new SXSSFExporterBuilder with the specified type and data.
     *
     * @param type               The class type of the data to be exported
     * @param data               The list of data objects to be exported
     * @param supplyExcelMaxRows The maximum number of rows supported by the Excel version
     */
    SXSSFExporterBuilder(
        Class<T> type,
        List<T> data,
        int supplyExcelMaxRows
    ) {
        this.type = type;
        this.data = data;
        this.supplyExcelMaxRows = supplyExcelMaxRows;
        this.maxRowsPerSheet = supplyExcelMaxRows;
        this.sheetMode = SheetMode.MULTI_SHEET;
        this.sheetName = null;
    }

    /**
     * Sets the sheet mode for the Excel exporter.
     *
     * @param sheetMode The mode to use for sheet management (ONE_SHEET or MULTI_SHEET)
     * @return This builder instance for method chaining
     */
    public SXSSFExporterBuilder<T> sheetMode(SheetMode sheetMode) {
        this.sheetMode = sheetMode;
        return this;
    }

    /**
     * Sets the maximum number of rows allowed per sheet.
     *
     * @param maxRowsPerSheet The maximum number of rows per sheet
     * @return This builder instance for method chaining
     * @throws ExcelException if maxRowsPerSheet exceeds the Excel version's maximum row limit
     */
    public SXSSFExporterBuilder<T> maxRows(int maxRowsPerSheet) {
        if (maxRowsPerSheet > supplyExcelMaxRows || maxRowsPerSheet < 2) {
            throw new ExcelException(String.format(
                "maxRowsPerSheet must be between 2 and %d (inclusive). Provided: %d.",
                supplyExcelMaxRows, maxRowsPerSheet));
        }
        this.maxRowsPerSheet = maxRowsPerSheet;
        return this;
    }

    /**
     * Sets the base name for sheets in the Excel file.
     * If set, each sheet will be named using this base name with an index suffix.
     *
     * <pre>
     * Example:
     * sheetName = "MySheet"
     * sheet = MySheet0, MySheet1, MySheet2, ...
     * </pre>
     *
     * @param sheetName The base name for sheets
     * @return This builder instance for method chaining
     */
    public SXSSFExporterBuilder<T> sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * Builds and returns a new SXSSFExporter instance with the configured settings.
     *
     * @return A new SXSSFExporter instance
     */
    public SXSSFExporter<T> build() {
        return new SXSSFExporter<T>(
            this.type,
            this.data,
            this.sheetMode,
            this.sheetName,
            this.maxRowsPerSheet
        );
    }
}
