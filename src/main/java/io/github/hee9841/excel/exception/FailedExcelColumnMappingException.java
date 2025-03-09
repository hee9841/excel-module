package io.github.hee9841.excel.exception;

public class FailedExcelColumnMappingException extends ExcelException {

    public FailedExcelColumnMappingException() {
        super();
    }

    public FailedExcelColumnMappingException(String message) {
        super(message);
    }

    public FailedExcelColumnMappingException(String message, Throwable cause) {
        super(message, cause);
    }

    public FailedExcelColumnMappingException(String message, String dtoTypeName) {
        super(message, dtoTypeName);
    }

    public FailedExcelColumnMappingException(String message, String dtoTypeName, Throwable cause) {
        super(message, dtoTypeName, cause);
    }
}
