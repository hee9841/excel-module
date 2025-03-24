package io.github.hee9841.excel.exception;


public class ExcelStyleException extends ExcelException {

    public ExcelStyleException() {
        super();
    }

    public ExcelStyleException(String message) {
        super(message);
    }

    public ExcelStyleException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelStyleException(String message, String dtoTypeName) {
        super(message, dtoTypeName);
    }

    public ExcelStyleException(String message, String dtoTypeName, Throwable cause) {
        super(message, dtoTypeName, cause);
    }
}
