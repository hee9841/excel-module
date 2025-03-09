package io.github.hee9841.excel.core;

import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.strategy.SheetStrategy;
import java.util.List;

public class ExcelExporterBuilder<T> {

    private final Class<T> type;
    private final List<T> data;

    private final int supplyExcelMaxRows;

    private int maxRowsPerSheet;
    private SheetStrategy sheetStrategy;
    private String sheetName;


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


    public ExcelExporterBuilder<T> sheetStrategy(SheetStrategy sheetStrategy) {
        this.sheetStrategy = sheetStrategy;
        return this;
    }

    public ExcelExporterBuilder<T> maxRows(int maxRowsPerSheet) {
        if (maxRowsPerSheet > supplyExcelMaxRows) {
            throw new ExcelException(String.format(
                "The maximum rows per sheet(%d) cannot exceed the supplied Excel sheet version's maximum row limit(%d).",
                maxRowsPerSheet, supplyExcelMaxRows));
        }
        this.maxRowsPerSheet = maxRowsPerSheet;
        return this;
    }


    public ExcelExporterBuilder<T> sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public ExcelExporter<T> build() {
        return new ExcelExporter<T>(
            this.type,
            this.data,
            this.sheetStrategy,
            this.sheetName,
            this.maxRowsPerSheet
        );
    }
}
