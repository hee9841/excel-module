package io.github.hee9841.excel.core;

import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.strategy.SheetStrategy;
import java.text.MessageFormat;
import java.util.List;

public class ExcelExporter<T> extends SXSSFExcelFile<T> {

    private static final String EXCEED_MAS_ROW_MSG_2ARGS =
        "The data size exceeds the maximum number of rows allowed per sheet. "
            + "The sheet strategy is set to ONE_SHEET but the data size ({0} rows) is larger than the maximum rows per sheet ({1} rows).\n"
            + "Please change the sheet strategy to MULTI_SHEET or reduce the data size.";

    private static final int ROW_START_INDEX = 0;
    private int currentRowIndex = ROW_START_INDEX;

    private SheetStrategy sheetStrategy;
    private final int maxRowsPerSheet;
    private final String sheetName;
    private int currentSheetIndex;


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


    public static <T> ExcelExporterBuilder<T> builder(Class<T> type, List<T> data) {
        return new ExcelExporterBuilder<>(type, data, supplyExcelVersion.getMaxRows());
    }

    @Override
    protected void validate(Class<?> type, List<T> data) {
        //아직 없음
    }

    @Override
    protected void createExcelFile(List<T> data) {
        // 1. If data is empty, create sheet and createHeader.
        if (data.isEmpty()) {
            createNewSheetWithHeader();
            logger.warn("Empty data provided - Excel file will be created with headers only.");
            return;
        }

        //2. Add rows
        createNewSheetWithHeader();
        addRows(data);
    }

    @Override
    public void addRows(List<T> data) {
        int leftDataSize = data.size();
        for (Object renderedData : data) {
            createBody(renderedData, currentRowIndex++);
            leftDataSize--;
            if (currentRowIndex == maxRowsPerSheet && leftDataSize > 0) {
                if (SheetStrategy.isOneSheet(sheetStrategy)) {
                    throw new ExcelException(
                        MessageFormat.format(EXCEED_MAS_ROW_MSG_2ARGS,
                            currentRowIndex + data.size(), maxRowsPerSheet), dtoTypeName);
                }
                //if multi sheet strategy
                createNewSheetWithHeader();
            }
        }
    }

    private void setSheetStrategy(SheetStrategy strategy, int dataSize) {
        // check sheet ...
        if (dataSize > maxRowsPerSheet && SheetStrategy.isOneSheet(strategy)) {
            throw new ExcelException(
                MessageFormat.format(EXCEED_MAS_ROW_MSG_2ARGS, dataSize, maxRowsPerSheet),
                dtoTypeName);
        }

        this.sheetStrategy = strategy;
        workbook.setZip64Mode(sheetStrategy.getZip64Mode());

        logger.debug("Set sheet strategy and Zip64Mode - strategy: {}, Zip64Mode:{}.",
            strategy.name(), sheetStrategy.getZip64Mode().name());
    }

    private void createNewSheetWithHeader() {
        currentRowIndex = ROW_START_INDEX;

        //sheet name 이 있으면 해당 sheet name + idx 으로 create
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
