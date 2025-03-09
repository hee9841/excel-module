package io.github.hee9841.excel.test.dto;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.meta.CellType;
import io.github.hee9841.excel.strategy.CellTypeStrategy;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

@Excel(
    cellTypeStrategy = CellTypeStrategy.AUTO
)
public class TypeAndFormatCheckForAutoDto {

    @ExcelColumn(headerName = "numberField")
    private int primitiveInt;

    @ExcelColumn(headerName = "numberField")
    private Integer wrapperInt;

    @ExcelColumn(headerName = "numberField")
    private double primitiveDouble;

    @ExcelColumn(headerName = "numberField")
    private Double wrapperDouble;

    @ExcelColumn(headerName = "numberField")
    private float primitiveFloat;

    @ExcelColumn(headerName = "numberField")
    private Float wrapperFloat;

    @ExcelColumn(headerName = "numberField")
    private long primitiveLong;

    @ExcelColumn(headerName = "numberField")
    private Long wrapperLong;

    @ExcelColumn(headerName = "numberField")
    private short primitiveShort;

    @ExcelColumn(headerName = "numberField")
    private Short wrapperShort;

    @ExcelColumn(headerName = "numberField")
    private byte primitiveByte;

    @ExcelColumn(headerName = "numberField")
    private Byte wrapperByte;


    @ExcelColumn(headerName = "stringField")
    private String stringType;

    @ExcelColumn(headerName = "boolField")
    private boolean primitiveBool;

    @ExcelColumn(headerName = "boolField")
    private Boolean wrapperBool;

    @ExcelColumn(headerName = "dateField")
    private Date date;

    @ExcelColumn(headerName = "localDateField")
    private LocalDate localDate;

    @ExcelColumn(headerName = "localDateTimeField")
    private LocalDateTime localDateTime;

    @ExcelColumn(headerName = "formalField", columnCellType = CellType.FORMULA)
    private String formal;

}
