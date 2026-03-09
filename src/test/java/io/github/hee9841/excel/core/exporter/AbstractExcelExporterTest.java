package io.github.hee9841.excel.core.exporter;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import io.github.hee9841.excel.exception.ExcelException;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("AbstractExcelExporter 테스트")
class AbstractExcelExporterTest {


    @DisplayName("워크북이 null이면 예외를 발생한다.")
    @Test
    void throwsExceptionWhenWorkbookIsNull() {
        assertThrows(ExcelException.class, () -> new TestExporter(null));
    }

    @DisplayName("시트 이름 프리픽스를 사용해 시트를 생성한다.")
    @Test
    void createSheetWithPrefix() throws IOException {
        //given
        String prefix = "MYSheet";
        int sheetIndex = 0;
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        //when
        Sheet sheet = testExporter.createSheet(prefix, sheetIndex);

        //then
        assertThat(sheet.getSheetName()).startsWith(prefix);
        assertEquals(sheetIndex, wb.getSheetIndex(sheet));

        wb.close();
    }

    @DisplayName("write에 null stream을 전달하면 IllegalArgumentException을 발생한다.")
    @Test
    void writeThrowsExceptionWhenStreamIsNull() {
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        assertThrows(IllegalArgumentException.class, () -> testExporter.write(null));
    }

    @DisplayName("워크북이 정상적으로 write 된다.")
    @Test
    void writeWorkbook() throws IOException {
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            testExporter.write(outputStream);
            assertThat(outputStream.toByteArray().length).isGreaterThan(0);
        }
    }

    @DisplayName("write() 호출 후 다시 write()를 호출하면 IllegalStateException을 발생한다.")
    @Test
    void writeAfterWriteThrowsException() throws IOException {
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            testExporter.write(outputStream);
        }

        assertThrows(IllegalStateException.class, () -> {
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                testExporter.write(outputStream);
            }
        });
    }

    @DisplayName("write() 호출 후 addRows()를 호출하면 IllegalStateException을 발생한다.")
    @Test
    void addRowsAfterWriteThrowsException() throws IOException {
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            testExporter.write(outputStream);
        }

        assertThrows(IllegalStateException.class, () -> testExporter.addRows(List.of()));
    }

    @DisplayName("close() 호출 후 write()를 호출하면 IllegalStateException을 발생한다.")
    @Test
    void writeAfterCloseThrowsException() {
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        testExporter.close();

        assertThrows(IllegalStateException.class, () -> {
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                testExporter.write(outputStream);
            }
        });
    }

    @DisplayName("close() 호출 후 addRows()를 호출하면 IllegalStateException을 발생한다.")
    @Test
    void addRowsAfterCloseThrowsException() {
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        testExporter.close();

        assertThrows(IllegalStateException.class, () -> testExporter.addRows(List.of()));
    }

    @DisplayName("close()는 여러 번 호출해도 예외가 발생하지 않는다.")
    @Test
    void closeIsIdempotent() {
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        testExporter.close();
        testExporter.close();
        testExporter.close();
        // 예외 없이 정상 종료
    }

    @DisplayName("try-with-resources로 사용할 수 있다.")
    @Test
    void canBeUsedWithTryWithResources() throws IOException {
        try (TestExporter testExporter = new TestExporter(new XSSFWorkbook());
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            testExporter.write(outputStream);
            assertThat(outputStream.toByteArray().length).isGreaterThan(0);
        }
        // 예외 없이 정상 종료
    }
    private static class TestExporter extends AbstractExcelExporter<T, Workbook> {

        private TestExporter(Workbook workbook) {
            super(workbook);
        }

        @Override
        protected void validate(Class<T> type, List<T> data) {
        }


        @Override
        protected void doAddRows(List<T> data) {
        }
    }
}
