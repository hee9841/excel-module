package io.github.hee9841.excel.meta;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import io.github.hee9841.excel.annotation.ExcelColumnStyle;
import io.github.hee9841.excel.exception.ExcelStyleException;
import io.github.hee9841.excel.style.CustomExcelCellStyle;
import io.github.hee9841.excel.style.ExcelCellStyle;
import io.github.hee9841.excel.style.align.CustomExcelAlign;
import io.github.hee9841.excel.style.align.DefaultExcelAlign;
import io.github.hee9841.excel.style.align.ExcelAlign;
import io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment;
import io.github.hee9841.excel.style.border.DefaultExcelBorder;
import io.github.hee9841.excel.style.border.ExcelBorder;
import io.github.hee9841.excel.style.border.ExcelBorderStyle;
import io.github.hee9841.excel.style.color.ColorPalette;
import io.github.hee9841.excel.style.color.ExcelColor;
import io.github.hee9841.excel.style.color.PaletteExcelColor;
import io.github.hee9841.excel.style.configurer.ExcelCellStyleConfigurer;
import java.lang.reflect.InvocationTargetException;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

public class CellStyleMappingTest {

    private Workbook wb;

    @BeforeEach
    void before() {
        wb = new SXSSFWorkbook();
    }

    @Nested
    class successStyleMappingTest {

        private CellStyle blackCenterThin;
        private CellStyle whiteGeneralCenterTopThin;

        @BeforeEach
        void before() {
            blackCenterThin = wb.createCellStyle();
            blackCenterThin.setFillForegroundColor(
                org.apache.poi.ss.usermodel.IndexedColors.BLACK.getIndex());
            blackCenterThin.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            blackCenterThin.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
            blackCenterThin.setVerticalAlignment(
                org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

            blackCenterThin.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            blackCenterThin.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            blackCenterThin.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            blackCenterThin.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);

            whiteGeneralCenterTopThin = wb.createCellStyle();
            whiteGeneralCenterTopThin.setFillForegroundColor(
                org.apache.poi.ss.usermodel.IndexedColors.WHITE.getIndex());
            whiteGeneralCenterTopThin.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            whiteGeneralCenterTopThin.setAlignment(
                org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL);
            whiteGeneralCenterTopThin.setVerticalAlignment(
                org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

            whiteGeneralCenterTopThin.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);

        }

        @DisplayName("enum cell style 경우 지정한 cell style이 적용되어야한다.(@ExcelColumn이 @Excel 보다 우선 적용됨)")
        @Test
        void enumCellStyle_applySuccess() {
            //given
            @Excel(
                defaultHeaderStyle = @ExcelColumnStyle(
                    cellStyleClass = TestEnumCellStyle.class,
                    enumName = "WHITE_GENERAL_CENTER_TOP_THIN"
                )
            )
            class TestDto {

                @ExcelColumn(headerName = "first",
                    bodyStyle = @ExcelColumnStyle(
                        cellStyleClass = TestEnumCellStyle.class,
                        enumName = "BLACK_CENTER_THIN"
                    )
                )
                private String firstFiled;

                @ExcelColumn(headerName = "second",
                    headerStyle = @ExcelColumnStyle(
                        cellStyleClass = TestEnumCellStyle.class,
                        enumName = "BLACK_CENTER_THIN"
                    )
                )
                private String secondFiled;
            }

            //when
            Map<Integer, ColumnInfo> resultMap = ColumnInfoMapper
                .of(TestDto.class, wb).map();

            //then
            assertCellStyleEquals(whiteGeneralCenterTopThin, resultMap.get(0).getHeaderStyle());
            assertCellStyleEquals(blackCenterThin, resultMap.get(0).getBodyStyle());

            assertCellStyleEquals(blackCenterThin, resultMap.get(1).getHeaderStyle());
            assertCellStyleEquals(wb.createCellStyle(), resultMap.get(1).getBodyStyle());
        }


        @DisplayName("class cell style 경우 지정한 cell style이 적용되어야한다.(@ExcelColumn이 @Excel 보다 우선 적용됨)")
        @Test
        void classCellStyle_applySuccess() {
            //given
            @Excel(
                defaultHeaderStyle = @ExcelColumnStyle(cellStyleClass = WhiteGeneralCenterTopThin.class)
            )
            class TestDto {

                @ExcelColumn(headerName = "first",
                    bodyStyle = @ExcelColumnStyle(cellStyleClass = BlackCenterThinStyle.class)
                )
                private String firstFiled;

                @ExcelColumn(headerName = "second",
                    headerStyle = @ExcelColumnStyle(cellStyleClass = BlackCenterThinStyle.class)
                )
                private String secondFiled;
            }

            //when
            Map<Integer, ColumnInfo> resultMap = ColumnInfoMapper
                .of(TestDto.class, wb).map();

            //then
            assertCellStyleEquals(whiteGeneralCenterTopThin, resultMap.get(0).getHeaderStyle());
            assertCellStyleEquals(blackCenterThin, resultMap.get(0).getBodyStyle());

            assertCellStyleEquals(blackCenterThin, resultMap.get(1).getHeaderStyle());
            assertCellStyleEquals(wb.createCellStyle(), resultMap.get(1).getBodyStyle());
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


    @DisplayName("CellStyle Mapping 예외 테스트")
    @Nested
    class CellStyleExceptionTest {

        @DisplayName("CellStyle이 Enum class일때, enum 값이 존재하지 않으면 예외 발생")
        @Test
        void enumClass_cannotFoundEnumValueName() {
            //given
            String expectedMsg = "Enum value 'NOTHING' not found in style class";
            //"Enum 관련 테스트 classes"
            @Excel(defaultHeaderStyle = @ExcelColumnStyle(
                cellStyleClass = TestEnumCellStyle.class,
                enumName = "NOTHING"
            ))
            class EnumClassStyleTestDTO {

                @ExcelColumn(headerName = "일")
                private String firstFiled;
            }
            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(EnumClassStyleTestDTO.class, wb);

            //when
            ExcelStyleException exception = assertThrows(
                ExcelStyleException.class,
                columnInfoMapper::map);

            //then
            assertTrue(exception.getMessage().contains(expectedMsg));
            assertInstanceOf(IllegalArgumentException.class, exception.getCause());
        }


        @DisplayName("CellStyle이 class일때, 기본 생성자가 존재하지 않으면 예외 발생")
        @Test
        void caseOfClass_NoSuchMethodException() {
            //given
            //NoConstructorException test classes
            @Excel(defaultHeaderStyle = @ExcelColumnStyle(
                cellStyleClass = NoConstructorStyleClass.class
            ))
            class NoConstructorTestDTO {

                @ExcelColumn(headerName = "일")
                private String firstFiled;
            }

            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(NoConstructorTestDTO.class, wb);

            //when
            ExcelStyleException exception = assertThrows(ExcelStyleException.class,
                columnInfoMapper::map);

            //then
            assertInstanceOf(NoSuchMethodException.class, exception.getCause());
        }


        @DisplayName("CellStyle이 class일때, 생성자가 private이면 예외 발생")
        @Test
        void caseOfClass_IllegalAccessException() {
            //given
            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(IllegalAccessTestDTO.class, wb);

            //when
            ExcelStyleException exception = assertThrows(ExcelStyleException.class,
                columnInfoMapper::map);

            //then
            assertInstanceOf(IllegalAccessException.class, exception.getCause());
        }

        @DisplayName("CellStyle이 class일때, 생성자가 내부에서 예외가 발생할 경우, 예외 발생")
        @Test
        void caseOfClass_InstantiationException() {
            //given
            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(InstantiationTestDTO.class, wb);

            //when
            ExcelStyleException exception = assertThrows(ExcelStyleException.class,
                columnInfoMapper::map);

            //then
            assertInstanceOf(InstantiationException.class, exception.getCause());
        }

        @DisplayName("CellStyle이 class일때, 클래스가 추상클래스 or 인터페이스이면 안됨")
        @Test
        void caseOfClass_InvocationTargetException() {
            //given
            ColumnInfoMapper columnInfoMapper = ColumnInfoMapper
                .of(InvocationTargetTestDTO.class, wb);

            //when
            ExcelStyleException exception = assertThrows(
                ExcelStyleException.class,
                columnInfoMapper::map);

            //then
            assertInstanceOf(InvocationTargetException.class, exception.getCause());
        }


        //IllegalAccessException test classes
        @Excel(defaultHeaderStyle = @ExcelColumnStyle(
            cellStyleClass = IllegalAccessStyleClass.class
        ))
        class IllegalAccessTestDTO {

            @ExcelColumn(headerName = "일")
            private String firstFiled;

        }


        //InstantiationException test classes
        @Excel(defaultHeaderStyle = @ExcelColumnStyle(
            cellStyleClass = InstantiationStyleClass.class
        ))
        class InstantiationTestDTO {

            @ExcelColumn(headerName = "일")
            private String firstFiled;

        }


        //InvocationTargetException test classes
        @Excel(defaultHeaderStyle = @ExcelColumnStyle(
            cellStyleClass = InvocationTargetStyleClass.class
        ))
        class InvocationTargetTestDTO {

            @ExcelColumn(headerName = "일")
            private String firstFiled;

        }

    }


    /**
     * test cell styles
     */

    enum TestEnumCellStyle implements ExcelCellStyle {
        BLACK_CENTER_THIN(
            PaletteExcelColor.of(ColorPalette.BLACK),
            CustomExcelAlign.from(ExcelHorizontalAlignment.HORIZONTAL_CENTER),
            DefaultExcelBorder.all(ExcelBorderStyle.THIN)
        ),
        WHITE_GENERAL_CENTER_TOP_THIN(
            PaletteExcelColor.of(ColorPalette.WHITE),
            DefaultExcelAlign.GENERAL_CENTER,
            DefaultExcelBorder.builder().top(ExcelBorderStyle.THIN).build()
        ),
        ;

        private final ExcelColor backgroundColor;
        private final ExcelAlign align;
        private final ExcelBorder border;

        TestEnumCellStyle(ExcelColor backgroundColor, ExcelAlign align, ExcelBorder border) {
            this.backgroundColor = backgroundColor;
            this.align = align;
            this.border = border;
        }

        @Override
        public void apply(CellStyle cellStyle) {
            backgroundColor.applyBackground(cellStyle);
            align.applyAlign(cellStyle);
            border.applyAllBorder(cellStyle);
        }
    }

    public static class BlackCenterThinStyle extends CustomExcelCellStyle {

        public BlackCenterThinStyle() {
            super();
        }

        @Override
        public void configure(ExcelCellStyleConfigurer configurer) {
            configurer.excelColor(PaletteExcelColor.of(ColorPalette.BLACK));
            configurer.excelAlign(
                CustomExcelAlign.from(ExcelHorizontalAlignment.HORIZONTAL_CENTER));
            configurer.excelBorder(DefaultExcelBorder.all(ExcelBorderStyle.THIN));
        }
    }

    public static class WhiteGeneralCenterTopThin extends CustomExcelCellStyle {

        public WhiteGeneralCenterTopThin() {
            super();
        }

        @Override
        public void configure(ExcelCellStyleConfigurer configurer) {
            configurer.excelColor(PaletteExcelColor.of(ColorPalette.WHITE));
            configurer.excelAlign(DefaultExcelAlign.GENERAL_CENTER);
            configurer.excelBorder(DefaultExcelBorder.builder().top(ExcelBorderStyle.THIN).build());
        }
    }


    public static class NoConstructorStyleClass implements ExcelCellStyle {

        private String tmp;

        public NoConstructorStyleClass(String name) {

        }

        @Override
        public void apply(CellStyle cellStyle) {

        }
    }

    public static class IllegalAccessStyleClass extends CustomExcelCellStyle {

        private IllegalAccessStyleClass() {
            super();
        }

        @Override
        public void configure(ExcelCellStyleConfigurer configurer) {

        }

        public static IllegalAccessStyleClass getInstance() {
            return new IllegalAccessStyleClass();
        }

        @Override
        public void apply(CellStyle cellStyle) {
        }

    }

    public static abstract class InstantiationStyleClass extends CustomExcelCellStyle {

        public InstantiationStyleClass() {
            super();
        }

        @Override
        public void apply(CellStyle cellStyle) {
        }

    }

    public static class InvocationTargetStyleClass extends CustomExcelCellStyle {

        public InvocationTargetStyleClass() {
            super();
            throw new RuntimeException("my Exception");
        }

        @Override
        public void configure(ExcelCellStyleConfigurer configurer) {

        }

        @Override
        public void apply(CellStyle cellStyle) {
        }
    }

}
