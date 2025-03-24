package io.github.hee9841.excel.style.color;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class IndexedExcelColorTest {

    @Test
    @DisplayName("applyBackground() 메소드는 CellStyle에 올바른 색상과 패턴을 적용해야 한다")
    void applyBackground_ShouldSetCorrectColorAndPattern() throws IOException {
        // given
        IndexedColors color = IndexedColors.BLUE;
        IndexedExcelColor excelColor = IndexedExcelColor.of(color);

        try(Workbook wb = new XSSFWorkbook()) {
            CellStyle cellStyle = wb.createCellStyle();

            // when
            excelColor.applyBackground(cellStyle);

            // then
            assertEquals(color.getIndex(), cellStyle.getFillForegroundColor());
            assertEquals(FillPatternType.SOLID_FOREGROUND, cellStyle.getFillPattern());
        }
    }
}
