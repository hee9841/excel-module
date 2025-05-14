package io.github.hee9841.excel.core.exporter;

import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.strategy.SheetStrategy;
import java.util.List;

/**
 * Builder class for creating and configuring {@link DefaultExcelExporter} instances.
 * This class implements the Builder pattern to provide a fluent interface for
 * configuring Excel export settings.
 *
 * <p>Default configuration:</p>
 * <ul>
 *     <li>Sheet Strategy: MULTI_SHEET</li>
 *     <li>Max Rows per Sheet: Excel 2007+ maximum - 1</li>
 *     <li>Sheet Name: null (default sheet names will be used)</li>
 * </ul>
 *
 * <p>Example usage:</p>
 * <pre>
 * DefaultExcelExporter&lt;MyData&gt; exporter = DefaultExcelExporter.builder(MyData.class, dataList)
 *     .sheetStrategy(SheetStrategy.ONE_SHEET)
 *     .maxRows(10000)
 *     .sheetName("MySheet")
 *     .build();
 * </pre>
 *
 * @param <T> The type of data to be exported
 */
public class ExcelExporterBuilder<T> {

    private final Class<T> type;
    private final List<T> data;

    private final int supplyExcelMaxRows;

    private int maxRowsPerSheet;
    private SheetStrategy sheetStrategy;
    private String sheetName;

    /**
     * Constructs a new ExcelExporterBuilder with the specified type and data.
     *
     * @param type               The class type of the data to be exported
     * @param data               The list of data objects to be exported
     * @param supplyExcelMaxRows The maximum number of rows supported by the Excel version
     */
    ExcelExporterBuilder(
        Class<T> type,
        List<T> data,
        int supplyExcelMaxRows
    ) {
        this.type = type;
        this.data = data;
        this.supplyExcelMaxRows = supplyExcelMaxRows;
        this.maxRowsPerSheet = supplyExcelMaxRows - 1;
        this.sheetStrategy = SheetStrategy.MULTI_SHEET;
        this.sheetName = null;
    }

    /**
     * Sets the sheet strategy for the Excel exporter.
     *
     * @param sheetStrategy The strategy to use for sheet management (ONE_SHEET or MULTI_SHEET)
     * @return This builder instance for method chaining
     */
    public ExcelExporterBuilder<T> sheetStrategy(SheetStrategy sheetStrategy) {
        this.sheetStrategy = sheetStrategy;
        return this;
    }

    /**
     * Sets the maximum number of rows allowed per sheet.
     *
     * @param maxRowsPerSheet The maximum number of rows per sheet
     * @return This builder instance for method chaining
     * @throws ExcelException if maxRowsPerSheet exceeds the Excel version's maximum row limit
     */
    public ExcelExporterBuilder<T> maxRows(int maxRowsPerSheet) {
        if (maxRowsPerSheet > supplyExcelMaxRows) {
            throw new ExcelException(String.format(
                "The maximum rows per sheet(%d) cannot exceed the supplied Excel sheet version's maximum row limit(%d).",
                maxRowsPerSheet, supplyExcelMaxRows));
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
    public ExcelExporterBuilder<T> sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * Builds and returns a new DefaultExcelExporter instance with the configured settings.
     *
     * @return A new DefaultExcelExporter instance
     */
    public DefaultExcelExporter<T> build() {
        return new DefaultExcelExporter<T>(
            this.type,
            this.data,
            this.sheetStrategy,
            this.sheetName,
            this.maxRowsPerSheet
        );
    }
}
