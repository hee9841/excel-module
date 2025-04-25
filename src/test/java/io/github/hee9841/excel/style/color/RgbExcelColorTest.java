package io.github.hee9841.excel.style.color;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import io.github.hee9841.excel.exception.ExcelStyleException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

class RgbExcelColorTest {

    Workbook wb;

    @BeforeEach
    void before() {
        wb = new XSSFWorkbook();
    }

    @AfterEach
    void after() throws IOException {
        wb.close();
    }

    @DisplayName("다양한 RGB 값 테스트")
    @ParameterizedTest
    @CsvSource({
        "0, 0, 0",
        "255, 255, 255",
        "128, 128, 128"
    })
    void testVariousRgbValues(int red, int green, int blue) {
        //given
        RgbExcelColor color = RgbExcelColor.rgb(red, green, blue);
        CellStyle cellStyle = wb.createCellStyle();

        //when
        color.applyBackground(cellStyle);

        //then
        XSSFColor appliedColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
        byte[] rgb = appliedColor.getRGB();

        assertEquals((byte) red, rgb[0]);
        assertEquals((byte) green, rgb[1]);
        assertEquals((byte) blue, rgb[2]);
    }

    @DisplayName("잘못된 RGB 값 테스트")
    @ParameterizedTest
    @CsvSource({
        "-1, 0, 0",
        "0, -1, 0",
        "0, 0, -1",
        "256, 128, 128",
        "128, 256, 128",
        "128, 128, 256",
    })
    void testInvalidRgbValues(int red, int green, int blue) {
        //given
        String expectedMsg = String.format("Invalid RGB values(r:%d, g:%d, b:%d). "
            + "Each value must be between 0 and 255.", red, green, blue);

        //when
        ExcelStyleException exception = assertThrows(ExcelStyleException.class, () -> {
            RgbExcelColor.rgb(red, green, blue);
        });

        //then
        assertEquals(expectedMsg, exception.getMessage());
    }
}
