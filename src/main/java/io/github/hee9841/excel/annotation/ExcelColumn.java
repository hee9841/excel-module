package io.github.hee9841.excel.annotation;

import io.github.hee9841.excel.meta.CellType;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    String headerName();

    int columnIndex() default -1;

    CellType columnCellType() default CellType._NONE;

    ExcelColumStyle headerStyle() default @ExcelColumStyle;

    ExcelColumStyle bodyStyle() default @ExcelColumStyle;
}
