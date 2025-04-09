package io.github.hee9841.excel.core;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import ch.qos.logback.classic.LoggerContext;
import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.annotation.ExcelColumnStyle;
import io.github.hee9841.excel.example.dto.TypeAutoDto;
import io.github.hee9841.excel.example.style.EnumCellStyleExample;
import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.format.CellFormats;
import io.github.hee9841.excel.strategy.CellTypeStrategy;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import io.github.hee9841.excel.strategy.DataFormatStrategy;
import io.github.hee9841.excel.strategy.SheetStrategy;
import io.github.hee9841.excel.style.CustomExcelCellStyle;
import io.github.hee9841.excel.style.align.DefaultExcelAlign;
import io.github.hee9841.excel.style.color.IndexedColors;
import io.github.hee9841.excel.style.color.IndexedExcelColor;
import io.github.hee9841.excel.style.configurer.ExcelCellStyleConfigurer;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import junit.log.MemoryAppender;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.slf4j.LoggerFactory;


@DisplayName("ExcelExporter 테스트")
class ExcelExporterByteOutputStreamTest {

    MemoryAppender memoryAppender;
    String loggerClassName;
    ByteArrayOutputStream os;

    @BeforeEach
    void beforeEach() {
        Logger logger = (Logger) LoggerFactory.getLogger(SXSSFExcelFile.class);
        loggerClassName = SXSSFExcelFile.class.getName();
        memoryAppender = new MemoryAppender();
        memoryAppender.setContext((LoggerContext) LoggerFactory.getILoggerFactory());
        logger.setLevel(Level.DEBUG);
        logger.addAppender(memoryAppender);
        memoryAppender.start();

        os = new ByteArrayOutputStream();

    }

    @AfterEach
    public void afterEach() throws IOException {
        memoryAppender.stop();
        os.close();
    }

    @DisplayName("data의 크기가 maxRows를 넘기고 ONE_SHEET 전략일 경우, build 시 오류를 발생한다.")
    @Test
    void dataSizeExceedMaxRowsThrowsException() {
        // given
        List<TestDto> data = new ArrayList<>();
        // Create more data items than the maxRows limit
        for (int i = 0; i < 11; i++) {
            data.add(new TestDto("test" + (i + 1), i + 1));
        }
        int maxRows = 10; // Set maxRows to 10, which is less than the data size

        // when & then
        ExcelException exception = assertThrows(ExcelException.class, () ->
            ExcelExporter.builder(TestDto.class, data)
                .maxRows(maxRows)
                .sheetStrategy(
                    SheetStrategy.ONE_SHEET) // Force ONE_SHEET strategy to ensure exception is thrown
                .build()
        );

        // Verify the exception message
        assertTrue(exception.getMessage()
            .contains("The data size exceeds the maximum number of rows allowed per sheet"));
    }


    @Test
    @DisplayName("엑셀 시트 버전의 최대 행을 넘는 max row 값을 설정할 수 없다.")
    void cannotExceedMaxRowOfImplementation() {
        //given
        List<TestDto> data = new ArrayList<>();
        int maxRowOfExcel2007 = 0x100000; // Excel 2007 최대 행 수 초과

        // when
        ExcelException exception = assertThrows(ExcelException.class, () -> ExcelExporter
            .builder(TestDto.class, data)
            .maxRows(maxRowOfExcel2007 + 1)
            .build());

        //then
        assertTrue(exception.getMessage()
            .contains("cannot exceed the supplied Excel sheet version's maximum row"));
    }

    @Test
    @DisplayName("엑셀 파일 생성 성공 시, log가 순서대로 생성되어야한다.")
    void validateLogSequence() throws IOException {
        // given
        List<TestDto> data = new ArrayList<>();
        data.add(new TestDto("test1", 1));
        data.add(new TestDto("test2", 2));

        ExcelExporter<TestDto> exporter = ExcelExporter
            .builder(TestDto.class, data).build();

        exporter.write(os);

        //then
        assertEquals(8, memoryAppender.getSize());
        assertTrue(memoryAppender.isPresent(0,
            "Set sheet strategy and Zip64Mode - strategy: MULTI_SHEET, Zip64Mode: Always.",
            Level.DEBUG));
        assertTrue(memoryAppender.isPresent(1, "Initializing", Level.INFO));
        assertTrue(memoryAppender.isPresent(2, "Mapping", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(3, "Create new Sheet", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(4, "Add rows data - row:1", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(5, "Add rows data - row:2", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(6, "Start to write Excel file", Level.INFO));
        assertTrue(memoryAppender.isPresent(7, "Successfully wrote Excel", Level.INFO));
    }


    @Test
    @DisplayName("빈 데이터로 엑셀 파일 생성시, 해더만 있는 빈 엑셀 파일을 생성한다.")
    void createEmptyExcelFile() throws IOException {
        // given
        List<TestDto> emptyData = new ArrayList<>();

        ExcelExporter<TestDto> exporter = ExcelExporter
            .builder(TestDto.class, emptyData)
            .build();

        // when & then
        exporter.write(os);

        try (Workbook workbook = WorkbookFactory.create(
            new ByteArrayInputStream(os.toByteArray()))) {
            assertNotNull(workbook);
            assertEquals(1, workbook.getNumberOfSheets()); // 시트 수 확인
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            assertEquals("이름", headerRow.getCell(0).getStringCellValue());
            assertEquals("번호", headerRow.getCell(1).getStringCellValue());
            assertNull(sheet.getRow(1));
        }

        assertEquals(7, memoryAppender.countEventsForLogger(loggerClassName));
        assertTrue(memoryAppender.isPresent("Empty data provided", Level.WARN));
    }

    @DisplayName("Sheet 관련 테스트")
    @Nested
    class SheetTest {

        @Test
        @DisplayName("multi sheet: 데이터 크기가 최대행을 넘을 때 자동으로 다음 시트를 생성한다.")
        void multiSheetMaxRowsExceedTest() throws IOException {
            // given
            int rowCnt = 18;
            List<TestDto> testData = new ArrayList<>();
            for (int i = 0; i < rowCnt; i++) {
                testData.add(new TestDto("test" + (i + 1), i + 1));
            }
            // when
            ExcelExporter.builder(TestDto.class, testData)
                .maxRows(10)
                .build().write(os);

            // then
            try (Workbook workbook = WorkbookFactory.create(
                new ByteArrayInputStream(os.toByteArray()))) {
                assertNotNull(workbook);
                assertEquals(2, workbook.getNumberOfSheets()); // 시트 수 확인
                for (int i = 0; i < 2; i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    //header 값 확인
                    Row headerRow = sheet.getRow(0);
                    assertEquals("이름", headerRow.getCell(0).getStringCellValue());
                    assertEquals("번호", headerRow.getCell(1).getStringCellValue());

                    // row 경계 값 null 확인
                    assertNotNull(sheet.getRow(1));
                    assertNotNull(sheet.getRow(9));
                    assertNull(sheet.getRow(10));
                }
            }

            assertTrue(memoryAppender
                .isPresent(
                    "Set sheet strategy and Zip64Mode - strategy: MULTI_SHEET, Zip64Mode: Always",
                    Level.DEBUG)
            );
            assertEquals(rowCnt, memoryAppender.search("Add rows data", Level.DEBUG).size());
            assertEquals(2, memoryAppender.search("Create new Sheet", Level.DEBUG).size());

        }

        @Test
        @DisplayName("one sheet: 데이터가 최대 행 초과 시 예외을 발생한다.")
        void throwExceptionWhenOneSheetAndExceedMaxRows() {
            // given
            List<TestDto> testData = new ArrayList<>();
            for (int i = 0; i < 18; i++) {
                testData.add(new TestDto("test" + (i + 1), i + 1));
            }

            //when
            ExcelException exception = assertThrows(ExcelException.class,
                () -> ExcelExporter.builder(TestDto.class, testData)
                    .maxRows(10)
                    .sheetStrategy(SheetStrategy.ONE_SHEET)
                    .build()
            );

            //then
            assertTrue(exception.getMessage()
                .contains("The data size exceeds the maximum number of rows allowed per sheet."));

            //
            assertEquals(2, memoryAppender.getSize());
            assertTrue(memoryAppender.isPresent("Initializing", Level.INFO));
            assertTrue(memoryAppender.isPresent(
                "Set sheet strategy and Zip64Mode - strategy: ONE_SHEET, Zip64Mode: AsNeeded.",
                Level.DEBUG));
        }

        @Test
        @DisplayName("시트 이름 지정 시 지정한 이름으로 시트 생성한다.")
        void createSheetWithSpecifiedName() throws IOException {
            //given
            List<TestDto> testData = new ArrayList<>();
            for (int i = 0; i < 18; i++) {
                testData.add(new TestDto("test" + (i + 1), i + 1));
            }

            //when
            ExcelExporter.builder(TestDto.class, testData)
                .sheetName("TestSheet")
                .maxRows(10)
                .build()
                .write(os);

            //then
            try (Workbook workbook = WorkbookFactory.create(
                new ByteArrayInputStream(os.toByteArray()))) {
                assertEquals("TestSheet0", workbook.getSheetAt(0).getSheetName());
                assertEquals("TestSheet1", workbook.getSheetAt(1).getSheetName());
            }

            assertTrue(memoryAppender.isPresent("Create new Sheet : TestSheet0.", Level.DEBUG));
            assertTrue(memoryAppender.isPresent("Create new Sheet : TestSheet1.", Level.DEBUG));

        }
    }


    @Test
    @DisplayName("Formula 타입인 cell는 함수 값이 적용 되어야한다.")
    void checkFormulaType() throws IOException {
        List<TypeAutoDto> testData = new ArrayList<>();
        int sum = 0;
        for (int i = 1; i < 11; i++) {
            TypeAutoDto testDto = new TypeAutoDto();
            testDto.setPrimitiveInt(i);
            if (i == 10) {
                testDto.setFormal("SUM(A2:A11)");
            }
            sum += i;
            testData.add(testDto);
        }

        //when
        ExcelExporter.builder(TypeAutoDto.class, testData)
            .sheetName("TestSheet")
            .build()
            .write(os);

        //  then
        try (Workbook workbook = WorkbookFactory
            .create(new ByteArrayInputStream(os.toByteArray()))) {
            Sheet sheet = workbook.getSheet("TestSheet0");

            for (int i = 1; i < 10; i++) {
                Row row = sheet.getRow(i);
                assertEquals(i, row.getCell(0).getNumericCellValue());
                assertEquals("", row.getCell(18).getStringCellValue());
            }
            Row lastRow = sheet.getRow(10);

            Cell formulaCell = lastRow.getCell(18);

            // Formula 확인
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(formulaCell);

            assertEquals(10, lastRow.getCell(0).getNumericCellValue());
            assertEquals(sum, formulaCell.getNumericCellValue());
        }
    }

    @DisplayName("enum 타입인 cell은 toString값이 적용되어야 한다.")
    @Test
    void checkEnumType() throws IOException {
        @Excel
        class TestDto {

            @ExcelColumn(headerName = "gender")
            final
            Gender gender;

            TestDto() {
                this.gender = Gender.MALE;
            }
        }
        List<TestDto> testData = new ArrayList<>();
        testData.add(new TestDto());

        //when
        ExcelExporter.builder(TestDto.class, testData)
            .build()
            .write(os);

        //  then
        try (Workbook workbook = WorkbookFactory
            .create(new ByteArrayInputStream(os.toByteArray()))) {
            Cell cell = workbook.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("M", cell.getStringCellValue());
        }
    }

    @DisplayName("cell style 테스트")
    @Nested
    class CellStyleTest {

        @DisplayName("지정한 cell의 style에 맞게 엑셀 파일이 생성 된다.")
        @Test
        void createExcelFileWithSpecifiedCellStyle() throws IOException {
            //given
            SXSSFWorkbook wb = new SXSSFWorkbook();
            // default header : GREY_25_PERCENT, CENTER_CENTER, ALL_BORDER_THICK
            CellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerStyle.setBorderTop(BorderStyle.THICK);
            headerStyle.setBorderBottom(BorderStyle.THICK);
            headerStyle.setBorderLeft(BorderStyle.THICK);
            headerStyle.setBorderRight(BorderStyle.THICK);

            //red header : RED, CENTER_CENTER, ALL_BORDER_THICK
            CellStyle redHeaderStyle = wb.createCellStyle();
            redHeaderStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            redHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            redHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
            redHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            redHeaderStyle.setBorderTop(BorderStyle.THICK);
            redHeaderStyle.setBorderBottom(BorderStyle.THICK);
            redHeaderStyle.setBorderLeft(BorderStyle.THICK);
            redHeaderStyle.setBorderRight(BorderStyle.THICK);

            //DefaultBodyStyle : WHITE, LEFT_BOTTOM
            CellStyle bodyStyle = wb.createCellStyle();
            bodyStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            bodyStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            bodyStyle.setAlignment(HorizontalAlignment.LEFT);
            bodyStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

            wb.close();

            List<TestDto> testData = new ArrayList<>();
            testData.add(new TestDto("name value", 1));

            //when
            ExcelExporter.builder(TestDto.class, testData)
                .build()
                .write(os);

            //  then
            try (Workbook workbook = WorkbookFactory
                .create(new ByteArrayInputStream(os.toByteArray()))) {

                Sheet sheet = workbook.getSheetAt(0);

                //header
                Row headers = sheet.getRow(0);
                assertCellStyleEquals(headerStyle, headers.getCell(0).getCellStyle());
                assertCellStyleEquals(redHeaderStyle, headers.getCell(1).getCellStyle());

                //body
                Row bodyRow = sheet.getRow(1);
                assertCellStyleEquals(bodyStyle, bodyRow.getCell(0).getCellStyle());
                assertCellStyleEquals(bodyStyle, bodyRow.getCell(1).getCellStyle());

            }


        }

        private void assertCellStyleEquals(CellStyle expected, CellStyle actual) {
            //color
            assertEquals(expected.getFillForegroundColor(), actual.getFillForegroundColor());
            assertEquals(expected.getFillPattern(), actual.getFillPattern());

            //align
            assertEquals(expected.getAlignment(), actual.getAlignment());
            assertEquals(expected.getVerticalAlignment(), actual.getVerticalAlignment());

            //border
            assertEquals(expected.getBorderTop(), actual.getBorderTop());
            assertEquals(expected.getBorderBottom(), actual.getBorderBottom());
            assertEquals(expected.getBorderLeft(), actual.getBorderLeft());
            assertEquals(expected.getBorderRight(), actual.getBorderRight());
        }
    }

    @DisplayName("지정한 dataformat이 적용되엇 엑셀 파일이 생성 되었는지 확인한다.")
    @Test
    void createExcelFileWithSpecifiedDataFormat() throws IOException {
        //given
        List<TestDto> testData = new ArrayList<>();
        testData.add(new TestDto("name value", 1));

        String signUpDatePattern = CellFormats.DEFAULT_DATE_FORMAT;

        //when
        ExcelExporter.builder(TestDto.class, testData)
            .build()
            .write(os);

        //  then
        try (Workbook workbook = WorkbookFactory
            .create(new ByteArrayInputStream(os.toByteArray()))) {
            Sheet sheet = workbook.getSheetAt(0);

            Row row1 = sheet.getRow(1);
            Cell signUpDateCell = row1.getCell(2);
            String actualFormat = signUpDateCell.getCellStyle().getDataFormatString();

            assertEquals(signUpDatePattern, actualFormat);
        }
    }


    @Excel(
        columnIndexStrategy = ColumnIndexStrategy.USER_DEFINED,
        defaultHeaderStyle = @ExcelColumnStyle(
            cellStyleClass = EnumCellStyleExample.class,
            enumName = "GREY_25_PERCENT_CENTER_CENTER_ALL_BORDER_THICK"
        ),
        defaultBodyStyle = @ExcelColumnStyle(cellStyleClass = DefaultBodyStyle.class),
        cellTypeStrategy = CellTypeStrategy.AUTO,
        dataFormatStrategy = DataFormatStrategy.AUTO_BY_CELL_TYPE
    )
    static class TestDto {

        @ExcelColumn(headerName = "이름", columnIndex = 0)
        private final String name;

        @ExcelColumn(headerName = "번호", columnIndex = 1,
            headerStyle = @ExcelColumnStyle(
                cellStyleClass = EnumCellStyleExample.class,
                enumName = "RED_CENTER_CENTER_ALL_BORDER_THICK"
            )
        )
        private final int number;

        @ExcelColumn(headerName = "가입일", columnIndex = 2, format = CellFormats.DEFAULT_DATE_FORMAT)
        private final LocalDateTime signUpDate;

        public TestDto(String name, int number) {
            this.name = name;
            this.number = number;
            this.signUpDate = LocalDateTime.now();
        }
    }

    public static class DefaultBodyStyle extends CustomExcelCellStyle {

        public DefaultBodyStyle() {
            super();
        }

        @Override
        public void configure(ExcelCellStyleConfigurer configurer) {
            configurer.excelColor(IndexedExcelColor.of(IndexedColors.WHITE));
            configurer.excelAlign(DefaultExcelAlign.LEFT_BOTTOM);
        }
    }

    enum Gender {
        FEMALE("F"),
        MALE("M"),
        ;
        private final String gender;

        Gender(String gender) {
            this.gender = gender;
        }


        @Override
        public String toString() {
            return this.gender;
        }
    }

}

