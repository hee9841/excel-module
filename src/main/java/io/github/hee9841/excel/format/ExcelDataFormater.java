package io.github.hee9841.excel.format;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;

/**
 * A formatter class for applying format patterns to Excel cell styles.
 * Uses Apache POI's DataFormat for applying formatting patterns to cell styles.
 */
public class ExcelDataFormater {

    /**
     * The format pattern string
     */
    private final String pattern;
    /**
     * The Apache POI DataFormat object
     */
    private final DataFormat dataFormat;


    private ExcelDataFormater(DataFormat dataFormat, String pattern) {
        this.pattern = pattern;
        this.dataFormat = dataFormat;
    }

    /**
     * Factory method to create a new ExcelDataFormater instance.
     *
     * @param dataFormat The Apache POI DataFormat object
     * @param pattern    The format pattern string to apply
     * @return A new ExcelDataFormater instance
     */
    public static ExcelDataFormater of(DataFormat dataFormat, String pattern) {
        return new ExcelDataFormater(dataFormat, pattern);
    }

    /**
     * Applies the format pattern to the given cell style.
     * If the pattern is empty, null, or the default format, no formatting is applied.
     *
     * @param cellStyle The Apache POI CellStyle to apply the format to
     */
    public void apply(CellStyle cellStyle) {
        if (CellFormats.isNone(pattern)) {
            return;
        }

        cellStyle.setDataFormat(dataFormat.getFormat(pattern));
    }
}
