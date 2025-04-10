package io.github.hee9841.excel.style.color;


import io.github.hee9841.excel.exception.ExcelStyleException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Implementation of ExcelColor that uses RGB color values to apply
 * background colors to Excel cell styles.
 *
 * <p><strong>Note:</strong> This implementation currently only supports XSSF (XLSX) files
 * and will not work with HSSF (XLS) formats. Use {@link PaletteExcelColor} for HSSF
 * compatibility.</p>
 */
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

    /**
     * Creates an RgbExcelColor instance from the specified RGB values.
     * Each value must be between 0 and 255, inclusive.
     *
     * @param red   the red component (0-255)
     * @param green the green component (0-255)
     * @param blue  the blue component (0-255)
     * @return a new RgbExcelColor instance
     * @throws ExcelStyleException if any of the RGB values are outside the valid range
     */
    public static RgbExcelColor rgb(int red, int green, int blue) {
        if (red < MIN_RGB || red > MAX_RGB || green < MIN_RGB ||
            green > MAX_RGB || blue < MIN_RGB || blue > MAX_RGB) {
            throw new ExcelStyleException(String.format("Invalid RGB values(r:%d, g:%d, b:%d). "
                + "Each value must be between 0 and 255.", red, green, blue)
            );
        }

        return new RgbExcelColor((byte) red, (byte) green, (byte) blue);
    }

    /**
     * Applies the RGB color as a solid background to the provided cell style.
     * Creates an XSSFColor from the RGB components and sets it as the fill foreground color.
     *
     * <p><strong>Note:</strong> This method will only work with XSSFCellStyle objects.
     * Attempting to use this with HSSFCellStyle will have no effect or may cause errors.</p>
     *
     * @param cellStyle the cell style to which the background color will be applied,
     *                  must be an instance of XSSFCellStyle
     */
    @Override
    public void applyBackground(CellStyle cellStyle) {
        cellStyle.setFillForegroundColor(new XSSFColor(new byte[]{red, green, blue}));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
}
