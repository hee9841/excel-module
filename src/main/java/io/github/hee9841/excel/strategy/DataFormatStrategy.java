package io.github.hee9841.excel.strategy;

public enum DataFormatStrategy {
    NONE,
    //오직 Cell type이 저정되어 있거나 Auto(_NONE이 아닌 경우)에만 적용됨
    AUTO_BY_CELL_TYPE,
    ;

    public boolean isAutoByCellType() {
        return this == AUTO_BY_CELL_TYPE;
    }

    public boolean isNone() {
        return this == NONE;
    }
}
