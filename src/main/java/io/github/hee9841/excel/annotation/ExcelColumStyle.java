package io.github.hee9841.excel.annotation;

import io.github.hee9841.excel.style.ExcelCellStyle;
import io.github.hee9841.excel.style.NoCellStyle;
import java.lang.annotation.Target;

@Target({})
public @interface ExcelColumStyle {

    Class<? extends ExcelCellStyle> cellStyleClass() default NoCellStyle.class;

    String enumName() default "";
}
