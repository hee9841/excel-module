package io.github.hee9841.excel.format;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;

public class ExcelDataFormater {

    private final String pattern;
    private final DataFormat dataFormat;


    private ExcelDataFormater(DataFormat dataFormat, String pattern) {
        this.pattern = pattern;
        this.dataFormat = dataFormat;
    }

    public static ExcelDataFormater of(DataFormat dataFormat, String pattern) {
        return new ExcelDataFormater(dataFormat, pattern);
    }

    public void apply(CellStyle cellStyle) {
        if (CellFormats.isNone(pattern)) {
            return;
        }

        cellStyle.setDataFormat(dataFormat.getFormat(pattern));
    }
}
