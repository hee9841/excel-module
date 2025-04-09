package io.github.hee9841.excel.example.dto;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.meta.ColumnDataType;
import io.github.hee9841.excel.strategy.CellTypeStrategy;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

@Excel(
    columnIndexStrategy = ColumnIndexStrategy.USER_DEFINED,
    cellTypeStrategy = CellTypeStrategy.AUTO
)
public class TypeAutoDto {

    @ExcelColumn(headerName = "numberField", columnIndex = 0)
    private int primitiveInt;

    @ExcelColumn(headerName = "numberField", columnIndex = 1)
    private Integer wrapperInt;

    @ExcelColumn(headerName = "numberField", columnIndex = 2)
    private double primitiveDouble;

    @ExcelColumn(headerName = "numberField", columnIndex = 3)
    private Double wrapperDouble;

    @ExcelColumn(headerName = "numberField", columnIndex = 4)
    private float primitiveFloat;

    @ExcelColumn(headerName = "numberField", columnIndex = 5)
    private Float wrapperFloat;

    @ExcelColumn(headerName = "numberField", columnIndex = 6)
    private long primitiveLong;

    @ExcelColumn(headerName = "numberField", columnIndex = 7)
    private Long wrapperLong;

    @ExcelColumn(headerName = "numberField", columnIndex = 8)
    private short primitiveShort;

    @ExcelColumn(headerName = "numberField", columnIndex = 9)
    private Short wrapperShort;

    @ExcelColumn(headerName = "numberField", columnIndex = 10)
    private byte primitiveByte;

    @ExcelColumn(headerName = "numberField", columnIndex = 11)
    private Byte wrapperByte;

    @ExcelColumn(headerName = "stringField", columnIndex = 12)
    private String stringType;

    @ExcelColumn(headerName = "boolField", columnIndex = 13)
    private boolean primitiveBool;

    @ExcelColumn(headerName = "boolField", columnIndex = 14)
    private Boolean wrapperBool;

    @ExcelColumn(headerName = "dateField", columnIndex = 15)
    private Date date;

    @ExcelColumn(headerName = "localDateField", columnIndex = 16)
    private LocalDate localDate;

    @ExcelColumn(headerName = "localDateTimeField", columnIndex = 17)
    private LocalDateTime localDateTime;

    @ExcelColumn(headerName = "formalField", columnIndex = 18, columnCellType = ColumnDataType.FORMULA)
    private String formal;

    public void setPrimitiveInt(int primitiveInt) {
        this.primitiveInt = primitiveInt;
    }

    public void setFormal(String formal) {
        this.formal = formal;
    }

}
