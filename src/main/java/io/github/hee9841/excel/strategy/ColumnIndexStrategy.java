package io.github.hee9841.excel.strategy;

public enum ColumnIndexStrategy {
    FIELD_ORDER,
    USER_DEFINED,
    ;
    public boolean isUserDefined() {
        return this == USER_DEFINED;
    }

    public boolean isFieldOrder() {
        return this == FIELD_ORDER;
    }
}
