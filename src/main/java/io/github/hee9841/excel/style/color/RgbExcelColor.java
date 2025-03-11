package io.github.hee9841.excel.style.color;


import io.github.hee9841.excel.exception.ExcelStyleException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFColor;

//In current, only supports XSSFCellStyle
//HSSFCellStyle is not supported
public class RgbExcelColor implements ExcelColor {

    private static final int MIN_RGB = 0;
    private static final int MAX_RGB = 255;

    private final byte red;
    private final byte green;
    private final byte blue;

    private RgbExcelColor(byte red, byte green, byte blue) {
        this.red = red;
        this.green = green;
        this.blue = blue;
    }

    public static RgbExcelColor rgb(int red, int green, int blue) {
        if(red < MIN_RGB || red > MAX_RGB || green < MIN_RGB ||
            green > MAX_RGB || blue < MIN_RGB || blue > MAX_RGB) {
            throw new ExcelStyleException(String.format("Invalid RGB values(r:%d, g:%d, b:%d). "
                + "Each value must be between 0 and 255.", red, green, blue)
            );
        }

        return new RgbExcelColor((byte) red, (byte) green, (byte) blue);
    }

    @Override
    public void applyBackground(CellStyle cellStyle) {
        cellStyle.setFillForegroundColor(new XSSFColor(new byte[]{red, green, blue}));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
}
