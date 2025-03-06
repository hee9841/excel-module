package io.github.hee9841.excel.exception;

public class FailedExcelColumnMappingException extends ExcelException{
    private final Class<?> dtoType;

    public FailedExcelColumnMappingException(String message) {
        super(message);
        this.dtoType = null;
    }

    public FailedExcelColumnMappingException(String message, Class<?> dtoType) {
        super(String.format("%s (%s)", message, dtoType.getName()));
        this.dtoType = dtoType;
    }

    public FailedExcelColumnMappingException(String message,  Class<?> dtoType, Throwable cause) {
        super(String.format("%s (%s)", message, dtoType.getName()), cause);
        this.dtoType = dtoType;
    }
}
