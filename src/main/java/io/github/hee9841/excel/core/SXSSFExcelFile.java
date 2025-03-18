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

public abstract class SXSSFExcelFile<T> implements ExcelFile<T> {

    protected static final Logger logger = LoggerFactory.getLogger(SXSSFExcelFile.class);

    protected static final SpreadsheetVersion supplyExcelVersion = SpreadsheetVersion.EXCEL2007;

    protected SXSSFWorkbook workbook;
    protected Map<Integer, ColumnInfo> columnsMappingInfo;
    protected Sheet sheet;

    protected String dtoTypeName;


    protected SXSSFExcelFile() {
        this.workbook = new SXSSFWorkbook();
    }

    protected void initialize(Class<?> type, List<T> data) {
        this.dtoTypeName = type.getName();
        logger.info("Initialize to create Excel File for DTO class({}).", dtoTypeName);

        validate(type, data);

        logger.debug("Mapping DTO to Excel data - DTO class({}).", dtoTypeName);
        this.columnsMappingInfo = ColumnInfoMapper.of(type, workbook).map();
    }

    protected void validate(Class<?> type, List<T> data) {
    }

    protected abstract void createExcelFile(List<T> data);

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

    protected void createBody(Object data, int rowIndex) {
        logger.debug("Add rows data - row:{}.", rowIndex);
        Row row = sheet.createRow(rowIndex);
        for (Integer colIndex : columnsMappingInfo.keySet()) {
            ColumnInfo columnInfo = columnsMappingInfo.get(colIndex);
            try {
                Field field = FieldUtils.getField(data.getClass(), columnInfo.getFieldName(), true);
                Cell cell = row.createCell(colIndex);
                columnInfo.getColumnType().setCellValueByCellType(cell, field.get(data));
                cell.setCellStyle(columnInfo.getBodyStyle());
            } catch (IllegalAccessException e) {
                throw new ExcelException(
                    String.format("Failed to create body(column:%d, row:%d) : "
                            + "Access to field %s failed.",
                        colIndex, rowIndex, columnInfo.getFieldName()), e);
            }
        }
    }

    @Override
    public void write(OutputStream stream) throws IOException {
        logger.info("Start to write Excel file for DTO class({}).", dtoTypeName);
        try (SXSSFWorkbook autoCloseableWb = this.workbook;
            OutputStream autoCloseableStream = stream) {
            autoCloseableWb.write(autoCloseableStream);
        }
        logger.info("Successfully wrote Excel file for DTO class({}).", dtoTypeName);
    }
}
