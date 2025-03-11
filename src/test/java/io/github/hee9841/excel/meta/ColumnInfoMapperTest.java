package io.github.hee9841.excel.meta;

import static io.github.hee9841.excel.strategy.ColumnIndexStrategy.USER_DEFINED;
import static org.junit.jupiter.api.Assertions.*;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.exception.ExcelException;
import io.github.hee9841.excel.strategy.ColumnIndexStrategy;
import java.util.Map;
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

    @DisplayName("@Excel 어노테이션이 없으면, throw Exception ")
    @Test
    void noExcelAnnotation_throwException() {
        //given
        class NoExcelAnnotationDto {}

        String expectedMsg = "Missing the @Excel annotation.";

        ColumnInfoMapper columnInfoMapper = ColumnInfoMapper.of(
            NoExcelAnnotationDto.class, wb);

        //when
        String exceptionMsg = assertThrows(
            ExcelException.class,
            columnInfoMapper::map).getMessage();

        assertTrue(exceptionMsg.contains(expectedMsg));
    }

    @Nested
    @DisplayName("column index mapping 테스트")
    class ColumnIndexMappingTest{

        @DisplayName("FIELD_ORDER 전략일 경우, 필드 순서대로 맵핑된다.")
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
            Map<Integer, ColumnInfo> map = columnInfoMapper.map();


            //then
            ColumnInfo firstCol = map.get(0);
            assertEquals("firstHeader", firstCol.getHeaderName());
            assertEquals("firstField", firstCol.getFieldName());

            ColumnInfo secondCol = map.get(1);
            assertEquals("secondHeader", secondCol.getHeaderName());
            assertEquals("secondField", secondCol.getFieldName());
        }

        @DisplayName("USER_DEFINED 전략일 경우, 지정한 Index값으로 맵핑된다.")
        @Test
        void userDefined_MappedByUserDefined() {
            //given
            @Excel(columnIndexStrategy = USER_DEFINED)
            class TestExcelDto {
                @ExcelColumn(headerName = "firstHeader", columnIndex = 5)
                String firstField;

                @ExcelColumn(headerName = "secondHeader", columnIndex = 4)
                String secondField;
            }
            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(TestExcelDto.class, wb);

            //when
            Map<Integer, ColumnInfo> map = columnInfoMapper.map();


            //then
            ColumnInfo firstCol = map.get(5);
            assertEquals("firstHeader", firstCol.getHeaderName());
            assertEquals("firstField", firstCol.getFieldName());

            ColumnInfo secondCol = map.get(4);
            assertEquals("secondHeader", secondCol.getHeaderName());
            assertEquals("secondField", secondCol.getFieldName());
        }


        @DisplayName("FIELD_ORDER 전략일 경우, 지정한 columnIndex를 무시")
        @Test
        void fieldOrder_Ignore_columnIndexByUserDefined() {
            //given
            @Excel(columnIndexStrategy = ColumnIndexStrategy.FIELD_ORDER)
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
            Map<Integer, ColumnInfo> map = columnInfoMapper.map();

            //then
            assertEquals("일", map.get(0).getHeaderName());
            assertEquals("이", map.get(1).getHeaderName());
            assertEquals("삼", map.get(2).getHeaderName());
        }

        @Nested
        @DisplayName("column index mapping 예외")
        class ColumnIdxMappingException {


            @DisplayName("USER_DEFINED 전략일 경우, columnIndex를 지정하지 않는 경우 예외 발생")
            @Test
            void userDefined_AndNotSetColumnIndex_Failed_Mapping() {
                //given
                @Excel(columnIndexStrategy = USER_DEFINED)
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


            @DisplayName("USER_DEFINED 전략일 경우, columnIndex값이 음수일 때 예외 발생")
            @Test
            void userDefined_AndColumnIndexNegative_Failed_Mapping() {
                //given
                @Excel(columnIndexStrategy = USER_DEFINED)
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

            @DisplayName("USER_DEFINED 전략일 경우, columnIndex 값이 중복 되면 예외 발생")
            @Test
            void userDefined_DuplicateColumnIndex() {
                //given
                @Excel(columnIndexStrategy = USER_DEFINED)
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

}
