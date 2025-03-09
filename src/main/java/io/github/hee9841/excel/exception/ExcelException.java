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

    public ExcelException(String message, String dtoTypeName) {
        super(String.format("%s (%s)", message, dtoTypeName));

    }

    public ExcelException(String message, String dtoTypeName, Throwable cause) {
        super(String.format("%s (%s)", message, dtoTypeName), cause);
    }

}
