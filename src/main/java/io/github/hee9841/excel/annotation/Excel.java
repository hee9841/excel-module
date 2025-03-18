package io.github.hee9841.excel.annotation;

import io.github.hee9841.excel.strategy.CellTypeStrategy;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import io.github.hee9841.excel.strategy.DataFormatStrategy;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

    ColumnIndexStrategy columnIndexStrategy() default ColumnIndexStrategy.FIELD_ORDER;

    CellTypeStrategy cellTypeStrategy() default CellTypeStrategy.NONE;

    DataFormatStrategy dataFormatStrategy() default DataFormatStrategy.NONE;

    ExcelColumStyle defaultHeaderStyle() default @ExcelColumStyle;

    ExcelColumStyle defaultBodyStyle() default @ExcelColumStyle;
}
