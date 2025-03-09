package io.github.hee9841.excel.strategy;

public enum CellTypeStrategy {
    NONE,
    AUTO,
    ;

    public boolean isAuto() {
        return this == AUTO;
    }

    public boolean isNone() {
        return this == NONE;
    }
}
