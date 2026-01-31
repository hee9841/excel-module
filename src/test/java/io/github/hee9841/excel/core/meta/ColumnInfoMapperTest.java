package io.github.hee9841.excel.core.meta;

import static io.github.hee9841.excel.mode.ColumnIndexMode.USER_DEFINED;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.example.dto.TypeAutoDto;
import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.format.CellFormats;
import io.github.hee9841.excel.mode.CellTypeMode;
import io.github.hee9841.excel.mode.ColumnIndexMode;
import io.github.hee9841.excel.mode.DataFormatMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;


class ColumnInfoMapperTest {

    private Workbook wb;

    @BeforeEach
    void before() {
        wb = new SXSSFWorkbook();
    }

    @DisplayName("@Excel 어노테이션이 없으면 예외를 발생한다.")
    @Test
    void noExcelAnnotation_throwException() {
        //given
        class NoExcelAnnotationDto {

        }

        String expectedMsg = "Missing the @Excel annotation.";

        ColumnInfoMapper columnInfoMapper = ColumnInfoMapper.of(
            NoExcelAnnotationDto.class, wb);

        //when
        String exceptionMsg = assertThrows(
            ExcelException.class,
            columnInfoMapper::map).getMessage();

        assertTrue(exceptionMsg.contains(expectedMsg));
    }

    @DisplayName("@ExcelColumn 어노테이션이 하나라도 없으면 예외를 발생한다.")
    @Test
    void noExcelColumnAnnotation_throwException() {
        //given
        @Excel
        class TestDto {

            String id;
        }

        String expectedMsg = "No @ExcelColumn annotations found in class";

        ColumnInfoMapper columnInfoMapper = ColumnInfoMapper.of(
            TestDto.class, wb);

        //when
        String exceptionMsg = assertThrows(
            ExcelException.class,
            columnInfoMapper::map).getMessage();

        assertTrue(exceptionMsg.contains(expectedMsg));
    }

    @Nested
    class ColumnIndexMappingTest {

        @DisplayName("FIELD_ORDER 모드일 경우, 필드 순서대로 맵핑된다.")
        @Test
        void fieldOrder_MappedByFieldOrder() {
            //given
            @Excel
            class TestExcelDto {

                @ExcelColumn(headerName = "firstHeader")
                String firstField;

                @ExcelColumn(headerName = "secondHeader")
                String secondField;
            }
            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(TestExcelDto.class, wb);

            //when
            List<ColumnInfo> columnInfos = columnInfoMapper.map();

            //then
            ColumnInfo firstCol = findByIndex(columnInfos, 0);
            assertEquals("firstHeader", firstCol.getHeaderName());
            assertEquals("firstField", firstCol.getFieldName());

            ColumnInfo secondCol = findByIndex(columnInfos, 1);
            assertEquals("secondHeader", secondCol.getHeaderName());
            assertEquals("secondField", secondCol.getFieldName());
        }

        @DisplayName("USER_DEFINED 모드일 경우, 지정한 Index값으로 맵핑된다.")
        @Test
        void userDefined_MappedByUserDefined() {
            //given
            @Excel(columnIndexMode = USER_DEFINED)
            class TestExcelDto {

                @ExcelColumn(headerName = "firstHeader", columnIndex = 5)
                String firstField;

                @ExcelColumn(headerName = "secondHeader", columnIndex = 4)
                String secondField;
            }
            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(TestExcelDto.class, wb);

            //when
            List<ColumnInfo> columnInfos = columnInfoMapper.map();

            //then
            ColumnInfo firstCol = findByIndex(columnInfos, 5);
            assertEquals("firstHeader", firstCol.getHeaderName());
            assertEquals("firstField", firstCol.getFieldName());

            ColumnInfo secondCol = findByIndex(columnInfos, 4);
            assertEquals("secondHeader", secondCol.getHeaderName());
            assertEquals("secondField", secondCol.getFieldName());
        }


        @DisplayName("FIELD_ORDER 모드일 경우, 지정한 columnIndex를 무시")
        @Test
        void fieldOrder_Ignore_columnIndexByUserDefined() {
            //given
            @Excel(columnIndexMode = ColumnIndexMode.FIELD_ORDER)
            class TestExcelDto {

                @ExcelColumn(headerName = "일", columnIndex = 3)
                private String firstFiled;
                @ExcelColumn(headerName = "이", columnIndex = 6)
                private String secondFiled;
                @ExcelColumn(headerName = "삼", columnIndex = 1)
                private String thirdFiled;
            }

            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(TestExcelDto.class, wb);

            //when
            List<ColumnInfo> columnInfos = columnInfoMapper.map();

            //then
            assertEquals("일", findByIndex(columnInfos, 0).getHeaderName());
            assertEquals("이", findByIndex(columnInfos, 1).getHeaderName());
            assertEquals("삼", findByIndex(columnInfos, 2).getHeaderName());
        }

        @DisplayName("column index mapping 예외")
        @Nested
        class ColumnIdxMappingException {


            @DisplayName("USER_DEFINED 모드일 경우, columnIndex를 지정하지 않는 경우 예외 발생")
            @Test
            void userDefined_AndNotSetColumnIndex_Failed_Mapping() {
                //given
                @Excel(columnIndexMode = USER_DEFINED)
                class TestExcelDto {

                    @ExcelColumn(headerName = "일")
                    private String firstFiled;
                }
                String expectedMsg = "is negative";

                ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                    .of(TestExcelDto.class, wb);

                //when
                ExcelException exception = assertThrows(
                    ExcelException.class,
                    columnInfoMapper::map);

                //then
                assertTrue(exception.getMessage().contains(expectedMsg));

            }


            @DisplayName("USER_DEFINED 모드일 경우, columnIndex값이 음수일 때 예외 발생")
            @Test
            void userDefined_AndColumnIndexNegative_Failed_Mapping() {
                //given
                @Excel(columnIndexMode = USER_DEFINED)
                class TestExcelDto {

                    @ExcelColumn(headerName = "일", columnIndex = -5)
                    private String firstFiled;
                }

                String expectedMsg = "is negative";

                ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                    .of(TestExcelDto.class, wb);

                //when
                ExcelException exception = assertThrows(
                    ExcelException.class,
                    columnInfoMapper::map);

                //then
                assertTrue(exception.getMessage().contains(expectedMsg));

            }

            @DisplayName("USER_DEFINED 모드일 경우, columnIndex 값이 중복 되면 예외 발생")
            @Test
            void userDefined_DuplicateColumnIndex() {
                //given
                @Excel(columnIndexMode = USER_DEFINED)
                class TestExcelDto {

                    @ExcelColumn(headerName = "first", columnIndex = 3)
                    private String firstFiled;

                    @ExcelColumn(headerName = "second", columnIndex = 3)
                    private String secondFiled;
                }

                String expectedMsg = "Duplicate value";

                ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                    .of(TestExcelDto.class, wb);

                //when
                ExcelException exception = assertThrows(
                    ExcelException.class,
                    columnInfoMapper::map);

                //then
                assertTrue(exception.getMessage().contains(expectedMsg));

            }

        }
    }

    @Nested
    class ColumnDataTypeMappingTest {

        @DisplayName("필드의 타입과 맞지 않는 타입을 지정했을 경우, 예외 발생")
        @Test
        void throwException_WhenIncompatibleCellTypeSpecified() {
            //given
            @Excel
            class TestDto {

                @ExcelColumn(headerName = "first",
                    columnCellType = ColumnDataType.NUMBER
                )
                String firsField;
            }
            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper.of(TestDto.class, wb);

            //when
            ExcelException exception = assertThrows(
                ExcelException.class, columnInfoMapper::map);
            //then
            assertTrue(
                exception.getMessage().contains("Invalid cell type : The cell type of "));

            assertTrue(
                exception.getMessage().contains("is not compatible with the specified cell type"));

        }

        @DisplayName("AUTO: 필드 타입에 따라 자동으로 또는 사용자가 지정한 타입에 맞게 맵핑된다.")
        @Test
        void cellTypeModeIsAuto_applyAutoType() {
            //given && when
            List<ColumnInfo> columnInfos = ColumnInfoMapper
                .of(TypeAutoDto.class, wb).map();

            //then
            assertEquals("primitiveInt", findByIndex(columnInfos, 0).getFieldName());
            assertEquals("numberField", findByIndex(columnInfos, 0).getHeaderName());

            for (ColumnInfo columnInfo : columnInfos) {

                switch (columnInfo.getHeaderName()) {
                    case "numberField":
                        assertEquals(ColumnDataType.NUMBER, columnInfo.getColumnType());
                        break;
                    case "stringField":
                        assertEquals(ColumnDataType.STRING, columnInfo.getColumnType());
                        break;
                    case "boolField":
                        assertEquals(ColumnDataType.BOOLEAN, columnInfo.getColumnType());
                        break;
                    case "dateField":
                        assertEquals(ColumnDataType.DATE, columnInfo.getColumnType());
                        break;
                    case "localDateField":
                        assertEquals(ColumnDataType.LOCAL_DATE, columnInfo.getColumnType());
                        break;
                    case "localDateTimeField":
                        assertEquals(ColumnDataType.LOCAL_DATE_TIME, columnInfo.getColumnType());
                        break;
                }
            }
        }

        @DisplayName("cellType 모드와 columnCellType을 지정하지 않았을 경우, None type으로 지정한다.")
        @Test
        void cellTypeModeAndColumnType_IsNone_applyNoneType() {
            //given
            @Excel
            class TestDto {

                @ExcelColumn(headerName = "first")
                Integer firsField;
            }

            //when
            List<ColumnInfo> columnInfos = ColumnInfoMapper
                .of(TestDto.class, wb).map();

            //then
            assertEquals(ColumnDataType._NONE, findByIndex(columnInfos, 0).getColumnType());
        }
    }


    @Nested
    class DataFormatMappingTest {

        @DisplayName("dataFormat모드가 None: 각 컬럼에 format을 지정 안하면 none으로, 지정한면 지정한 패턴으로 적용된다.")
        @Test
        void modeIsNone_appliedByPattern() {
            //given
            @Excel
            class TestDto {

                //1. _NONE (default) -> general
                @ExcelColumn(headerName = "noneFormat")
                private int noneInt;

                @ExcelColumn(headerName = "noneFormat")
                private String noneString;

                @ExcelColumn(headerName = "noneFormat")
                private boolean noneBoolean;

                @ExcelColumn(headerName = "noneFormat")
                private Date noneDate;

                @ExcelColumn(headerName = "noneFormat")
                private LocalDate noneLocalDate;

                @ExcelColumn(headerName = "noneFormat")
                private LocalDateTime noneLocalDateTime;

                //2. 지정하면 -> 지정한 값으로
                @ExcelColumn(headerName = "won",
                    format = CellFormats.KR_WON_FORMAT
                )
                private Integer wonDataFormat;

                @ExcelColumn(headerName = "dateTImeFormat",
                    format = CellFormats.DEFAULT_DATE_TIME_FORMAT
                )
                private LocalDateTime localDateTime;
            }

            // && when
            List<ColumnInfo> columnInfos = ColumnInfoMapper
                .of(TestDto.class, wb).map();

            //then
            for (ColumnInfo columnInfo : columnInfos) {
                String formatString = columnInfo.getBodyStyle().getDataFormatString();
                assertFormatByHeaderName(columnInfo.getHeaderName(), formatString);

            }
        }


        @DisplayName("dataFormatMode: AUTO_BY_CELL_TYPE : 지정한 pattern에 따라 적용되거나, celltype이 지정되면 cellType에 따라 적용된다.")
        @Test
        void modeAutoByCellType_AndNoneCellType() {
            @Excel(dataFormatMode = DataFormatMode.AUTO_BY_CELL_TYPE)
            class TestDto {

                //1. cell type is none and format is none -> is none(general)
                @ExcelColumn(headerName = "noneFormat")
                private int noneInt;

                @ExcelColumn(headerName = "noneFormat")
                private LocalDate noneLocalDate;

                @ExcelColumn(headerName = "noneFormat")
                private LocalDateTime noneLocalDateTime;


                //2. cell type is none and format is specific -> 지정한 값으로
                @ExcelColumn(headerName = "dateFormat", format = CellFormats.DEFAULT_DATE_FORMAT)
                private LocalDate noneSpecificLocalDate;

                @ExcelColumn(headerName = "dateTImeFormat", format = CellFormats.DEFAULT_DATE_TIME_FORMAT)
                private LocalDateTime noneSpecificLocalDateTime;

                @ExcelColumn(headerName = "won",
                    format = CellFormats.KR_WON_FORMAT
                )
                private Integer wonDataFormat1;


                //3. cell type is auto and format is none -> cell type의 값으로
                @ExcelColumn(headerName = "dateFormat", columnCellType = ColumnDataType.AUTO)
                private LocalDate cellAutoNoneFormatLocalDate;

                @ExcelColumn(headerName = "dateTImeFormat", columnCellType = ColumnDataType.AUTO)
                private LocalDateTime cellAutoNoneFormatLocalDateTime;


                //4. cell type is auto and format is specific -> 지정한 값으로
                @ExcelColumn(headerName = "yyyy.MM.dd", columnCellType = ColumnDataType.AUTO,
                    format = "yyyy.MM.dd"
                )
                private LocalDate cellAutoSpecificFormatLocalDate;

                @ExcelColumn(headerName = "MM.dd.yy(HH:mm:ss)", columnCellType = ColumnDataType.AUTO,
                    format = "MM.dd.yy(HH:mm:ss)"
                )
                private LocalDateTime cellAutoSpecificFormatLocalDateTime;

                @ExcelColumn(headerName = "won", columnCellType = ColumnDataType.AUTO,
                    format = CellFormats.KR_WON_FORMAT
                )
                private Integer wonDataFormat2;

            }

            // && when
            List<ColumnInfo> columnInfos = ColumnInfoMapper
                .of(TestDto.class, wb).map();

            //then
            for (ColumnInfo columnInfo : columnInfos) {
                String formatString = columnInfo.getBodyStyle().getDataFormatString();
                assertFormatByHeaderName(columnInfo.getHeaderName(), formatString);

            }
        }

        @DisplayName("dataFormatMode: AUTO_BY_CELL_TYPE, CellType Auto : 기본으로 cellType에 따라 적용되고 지정하면 지정한 pattern이 적용된다.")
        @Test
        void modeAutoByCellType_AndCellTypeisAuto() {
            //givne
            @Excel(
                dataFormatMode = DataFormatMode.AUTO_BY_CELL_TYPE,
                cellTypeMode = CellTypeMode.AUTO
            )
            class TestDto {

                //1. cell type is none(오토로 지정이 됨) and format is none -> auto(cell type) 로
                @ExcelColumn(headerName = "noneFormat")
                private int noneInt;

                @ExcelColumn(headerName = "dateFormat")
                private LocalDate noneLocalDate;

                @ExcelColumn(headerName = "dateTImeFormat")
                private LocalDateTime noneLocalDateTime;


                //2. cell type is none(오토로 지정이 됨) and format is specific -> 지정한 값으로
                @ExcelColumn(headerName = "yyyy.MM.dd", format = "yyyy.MM.dd")
                private LocalDate cellAutoSpecificFormatLocalDate;

                @ExcelColumn(headerName = "MM.dd.yy(HH:mm:ss)", format = "MM.dd.yy(HH:mm:ss)")
                private LocalDateTime cellAutoSpecificFormatLocalDateTime;

                @ExcelColumn(headerName = "won", format = CellFormats.KR_WON_FORMAT)
                private Integer wonDataFormat;

            }

            // && when
            List<ColumnInfo> columnInfos = ColumnInfoMapper
                .of(TestDto.class, wb).map();

            //then
            for (ColumnInfo columnInfo : columnInfos) {
                String formatString = columnInfo.getBodyStyle().getDataFormatString();
                assertFormatByHeaderName(columnInfo.getHeaderName(), formatString);

            }
        }


        private void assertFormatByHeaderName(String headerName, String formatPattern) {
            switch (headerName) {
                case "dateFormat":
                    assertEquals(CellFormats.DEFAULT_DATE_FORMAT, formatPattern);
                    break;
                case "dateTImeFormat":
                    assertEquals(CellFormats.DEFAULT_DATE_TIME_FORMAT, formatPattern);
                    break;
                case "won":
                    assertEquals(CellFormats.KR_WON_FORMAT, formatPattern);
                    break;
                case "yyyy.MM.dd":
                    assertEquals("yyyy.MM.dd", formatPattern);
                    break;
                case "MM.dd.yy(HH:mm:ss)":
                    assertEquals("MM.dd.yy(HH:mm:ss)", formatPattern);
                    break;
                default:
                    assertEquals("General", formatPattern);
                    break;
            }
        }
    }

    @Nested
    @DisplayName("validateField 테스트")
    class ValidateFieldTest {

        @DisplayName("enum 타입 필드는 @ExcelColumn으로 사용할 수 있다.")
        @Test
        void enumFieldType_shouldBeAllowed() {
            // given
            @Excel(cellTypeMode = CellTypeMode.AUTO)
            class TestDto {

                @ExcelColumn(headerName = "status")
                private TestStatus status;
            }

            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper.of(TestDto.class, wb);

            // when
            List<ColumnInfo> columnInfos = columnInfoMapper.map();

            // then
            assertEquals(1, columnInfos.size());
            assertEquals("status", columnInfos.get(0).getFieldName());
            assertEquals(ColumnDataType.ENUM, columnInfos.get(0).getColumnType());
        }

        @DisplayName("primitive 타입 필드는 @ExcelColumn으로 사용할 수 있다.")
        @Test
        void primitiveFieldType_shouldBeAllowed() {
            // given
            @Excel(cellTypeMode = CellTypeMode.AUTO)
            class TestDto {

                @ExcelColumn(headerName = "intValue")
                private int intValue;

                @ExcelColumn(headerName = "boolValue")
                private boolean boolValue;

                @ExcelColumn(headerName = "charValue")
                private char charValue;
            }

            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper.of(TestDto.class, wb);

            // when
            List<ColumnInfo> columnInfos = columnInfoMapper.map();

            // then
            assertEquals(3, columnInfos.size());
            assertEquals(ColumnDataType.NUMBER, findByIndex(columnInfos, 0).getColumnType());
            assertEquals(ColumnDataType.BOOLEAN, findByIndex(columnInfos, 1).getColumnType());
            assertEquals(ColumnDataType.STRING, findByIndex(columnInfos, 2).getColumnType());
        }

        @DisplayName("배열 타입 필드는 @ExcelColumn으로 사용할 수 없다.")
        @Test
        void arrayFieldType_shouldThrowException() {
            // given
            @Excel
            class TestDto {

                @ExcelColumn(headerName = "values")
                private int[] values;
            }

            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper.of(TestDto.class, wb);

            // when & then
            ExcelException exception = assertThrows(
                ExcelException.class, columnInfoMapper::map);

            assertTrue(exception.getMessage().contains("cannot be applied to array type"));
        }

        @DisplayName("허용되지 않은 타입 필드는 @ExcelColumn으로 사용할 수 없다.")
        @Test
        void notAllowedFieldType_shouldThrowException() {
            // given
            @Excel
            class TestDto {

                @ExcelColumn(headerName = "list")
                private List<String> list;
            }

            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper.of(TestDto.class, wb);

            // when & then
            ExcelException exception = assertThrows(
                ExcelException.class, columnInfoMapper::map);

            assertTrue(exception.getMessage().contains("can only be applied to allowed types"));
        }

        enum TestStatus {
            ACTIVE, INACTIVE
        }
    }

    private ColumnInfo findByIndex(List<ColumnInfo> columnInfos, int index) {
        return columnInfos.stream()
            .filter(columnInfo -> columnInfo.getIndex() == index)
            .findFirst()
            .orElseThrow(() -> new AssertionError("Column index not found: " + index));
    }

}
