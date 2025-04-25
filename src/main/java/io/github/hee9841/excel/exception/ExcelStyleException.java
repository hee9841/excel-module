package io.github.hee9841.excel.exception;


public class ExcelStyleException extends ExcelException {

    public ExcelStyleException(String message) {
        super(message);
    }

    public ExcelStyleException(String message, String dtoTypeName, Throwable cause) {
        super(message, dtoTypeName, cause);
    }
}
