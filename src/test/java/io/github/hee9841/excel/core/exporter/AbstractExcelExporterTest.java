package io.github.hee9841.excel.core.exporter;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import io.github.hee9841.excel.exception.ExcelException;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
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

    @DisplayName("write에 null stream을 전달하면 NPE를 발생한다.")
    @Test
    void writeThrowsExceptionWhenStreamIsNull() {
        Workbook wb = new XSSFWorkbook();
        TestExporter testExporter = new TestExporter(wb);

        assertThrows(NullPointerException.class, () -> testExporter.write(null));
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

    private static class TestExporter extends AbstractExcelExporter<Object, Workbook> {

        private TestExporter(Workbook workbook) {
            super(workbook);
        }

        @Override
        protected void validate(Class<?> type, List<Object> data) {
        }

        @Override
        protected void createExcel(List<Object> data) {
        }

        @Override
        public void addRows(List<Object> data) {
        }
    }
}
