package io.github.hee9841.excel.core.exporter;

import static org.junit.jupiter.api.Assertions.assertEquals;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.strategy.CellTypeStrategy;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("ExcelExporter interface 테스트")
class ExcelExporterTest {

    @DisplayName("인터페이스 타입으로 addRows 후 write가 정상 동작한다.")
    @Test
    void addRowsAndWriteViaInterface() throws IOException {
        // given
        List<TestDto> initialData = List.of(
            new TestDto("alpha", 1),
            new TestDto("beta", 2)
        );

        ExcelExporter<TestDto> exporter = SXSSFExporter.builder(TestDto.class, initialData)
            .maxRows(5)
            .build();

        // when
        exporter.addRows(List.of(new TestDto("gamma", 3)));

        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            exporter.write(os);

            // then
            try (Workbook workbook = WorkbookFactory.create(
                new ByteArrayInputStream(os.toByteArray()))) {
                Sheet sheet = workbook.getSheetAt(0);
                assertEquals("name", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("number", sheet.getRow(0).getCell(1).getStringCellValue());

                assertEquals("alpha", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals(1, (int) sheet.getRow(1).getCell(1).getNumericCellValue());

                assertEquals("beta", sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals(2, (int) sheet.getRow(2).getCell(1).getNumericCellValue());

                assertEquals("gamma", sheet.getRow(3).getCell(0).getStringCellValue());
                assertEquals(3, (int) sheet.getRow(3).getCell(1).getNumericCellValue());
            }
        }
    }

    @Excel(
        columnIndexStrategy = ColumnIndexStrategy.USER_DEFINED,
        cellTypeStrategy = CellTypeStrategy.AUTO
    )
    static class TestDto {
        @ExcelColumn(headerName = "name", columnIndex = 0)
        private final String name;

        @ExcelColumn(headerName = "number", columnIndex = 1)
        private final int number;

        TestDto(String name, int number) {
            this.name = name;
            this.number = number;
        }
    }
}
