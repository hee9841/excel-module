package io.github.hee9841.excel.exception;

public class ExcelException extends RuntimeException {
    public ExcelException() {
        super();
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }
}
