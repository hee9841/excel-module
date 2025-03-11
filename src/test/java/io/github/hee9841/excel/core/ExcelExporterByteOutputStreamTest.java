package io.github.hee9841.excel.core;

import static org.junit.jupiter.api.Assertions.*;


import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import ch.qos.logback.classic.LoggerContext;
import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import io.github.hee9841.excel.strategy.SheetStrategy;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import junit.log.MemoryAppender;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.slf4j.LoggerFactory;


@DisplayName("ExcelExporter 테스트")
class ExcelExporterByteOutputStreamTest {

    MemoryAppender memoryAppender;
    String loggerClassName;

    @BeforeEach
    void beforeEach() {
        Logger logger = (Logger) LoggerFactory.getLogger(SXSSFExcelFile.class);
        loggerClassName = SXSSFExcelFile.class.getName();
        memoryAppender = new MemoryAppender();
        memoryAppender.setContext((LoggerContext) LoggerFactory.getILoggerFactory());
        logger.setLevel(Level.DEBUG);
        logger.addAppender(memoryAppender);
        memoryAppender.start();

    }

    @AfterEach
    public void afterEach() {
        memoryAppender.stop();
    }
    


    @Test
    @DisplayName("지정한 최대 행이 적용되는 엑셀 시트 버전의 최대 행을 넘을 경우")
    void cannotExceedMaxRowOfImplementation() {
        //given
        List<TestDto> data = new ArrayList<>();
        int maxRowOfExcel2007 = 0x100000; // Excel 2007 최대 행 수 초과

        // when
        ExcelException exception = assertThrows(ExcelException.class, () -> ExcelExporter
            .builder(TestDto.class, data)
            .maxRows(maxRowOfExcel2007 +1)
            .build());

        //then
        assertTrue(exception.getMessage().contains("cannot exceed the supplied Excel sheet version's maximum row"));
    }

    @Test
    @DisplayName("엑셀 파일 생성 성공 시, log 생성 순서 테스트")
    void validateLogSequence() throws IOException{
        // given
        List<TestDto> data = new ArrayList<>();
        data.add(new TestDto("test1", 1));
        data.add(new TestDto("test2", 2));

        ExcelExporter<TestDto> exporter = ExcelExporter
            .builder(TestDto.class, data).build();

        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            exporter.write(os);
        }


        //then
        assertEquals(8, memoryAppender.getSize());
        assertTrue(memoryAppender.isPresent(0, "Initialize", Level.INFO));
        assertTrue(memoryAppender.isPresent(1, "Mapping", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(2, "Set sheet strategy and Zip64Mode - strategy: MULTI_SHEET, Zip64Mode:Always.", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(3, "Create new Sheet", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(4, "Add rows data - row:1", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(5, "Add rows data - row:2", Level.DEBUG));
        assertTrue(memoryAppender.isPresent(6, "Start to write Excel file", Level.INFO));
        assertTrue(memoryAppender.isPresent(7, "Successfully wrote Excel", Level.INFO));
    }


    @Test
    @DisplayName("빈 데이터로 엑셀 파일 생성시, header만 생성")
    void createEmptyExcelFile() throws IOException {
        // given
        List<TestDto> emptyData = new ArrayList<>();

        ExcelExporter<TestDto> exporter = ExcelExporter
            .builder(TestDto.class, emptyData)
            .build();

        // when & then
        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            exporter.write(os);

            try (Workbook workbook =  WorkbookFactory.create(new ByteArrayInputStream(os.toByteArray()))) {
                assertNotNull(workbook);
                assertEquals(1, workbook.getNumberOfSheets()); // 시트 수 확인
                Sheet sheet = workbook.getSheetAt(0);
                Row headerRow = sheet.getRow(0);
                assertEquals("이름", headerRow.getCell(0).getStringCellValue());
                assertEquals("번호", headerRow.getCell(1).getStringCellValue());
                assertNull(sheet.getRow(1));
            }
        }

        assertEquals(7, memoryAppender.countEventsForLogger(loggerClassName));
        assertTrue(memoryAppender.isPresent("Empty data provided",Level.WARN));
    }

    @Test
    @DisplayName("multi sheet: 데이터 크기가 최대행을 넘을 때 자동으로 다음 시트를 생성한다.")
    void multiSheetMaxRowsExceedTest() throws IOException{
        // given
        int rowCnt = 18;
        List<TestDto> testData = new ArrayList<>();
        for (int i = 0; i < rowCnt; i++) {
            testData.add(new TestDto("test" + (i + 1) , i +1));
        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();


        // when
        ExcelExporter.builder(TestDto.class, testData)
            .maxRows(10)
            .build().write(os);


        // then
        try (Workbook workbook =  WorkbookFactory.create(new ByteArrayInputStream(os.toByteArray()))) {
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

        os.close();

        assertTrue(memoryAppender
            .isPresent("Set sheet strategy and Zip64Mode - strategy: MULTI_SHEET, Zip64Mode:Always",
                Level.DEBUG)
        );
        assertEquals(rowCnt, memoryAppender.search("Add rows data", Level.DEBUG).size());
        assertEquals(2, memoryAppender.search("Create new Sheet", Level.DEBUG).size());

    }

    @Test
    @DisplayName("one sheet: 데이터가 최대행 초과 시 예외 발생")
    void throwExceptionWhenOneSheetAndExceedMaxRows(){
        // given
        List<TestDto> testData = new ArrayList<>();
        for (int i = 0; i < 18; i++) {
            testData.add(new TestDto("test" + (i + 1) , i +1));
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
        assertTrue(memoryAppender.isPresent("Initialize", Level.INFO));
        assertTrue(memoryAppender.isPresent("Mapping DTO", Level.DEBUG));
    }
    
    @Test
    @DisplayName("시트 이름 지정 시 지정한 이름으로 시트 생성")
    void createSheetWithSpecifiedName() throws IOException{
        //given
        List<TestDto> testData = new ArrayList<>();
        for (int i = 0; i < 18; i++) {
            testData.add(new TestDto("test" + (i + 1) , i +1));
        }

        ByteArrayOutputStream os = new ByteArrayOutputStream();

        //when
        ExcelExporter.builder(TestDto.class, testData)
            .sheetName("TestSheet")
            .maxRows(10)
            .build()
            .write(os);

        //then
        try (Workbook workbook =  WorkbookFactory.create(new ByteArrayInputStream(os.toByteArray()))) {
            assertEquals("TestSheet0", workbook.getSheetAt(0).getSheetName());
            assertEquals("TestSheet1", workbook.getSheetAt(1).getSheetName());
        }
        os.close();


        assertTrue(memoryAppender.isPresent("Create new Sheet : TestSheet0.", Level.DEBUG));
        assertTrue(memoryAppender.isPresent("Create new Sheet : TestSheet1.", Level.DEBUG));

    }
    

    @Excel(columnIndexStrategy = ColumnIndexStrategy.USER_DEFINED)
    static class TestDto {
        @ExcelColumn(headerName = "이름", columnIndex = 0)
        private String name;
        
        @ExcelColumn(headerName = "번호", columnIndex = 1)
        private int number;
        
        public TestDto(String name, int number) {
            this.name = name;
            this.number = number;
        }
    }
}

