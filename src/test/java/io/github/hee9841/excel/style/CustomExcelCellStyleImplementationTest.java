package io.github.hee9841.excel.style;

import static io.github.hee9841.excel.style.align.alignment.HorizontalAlignment.HORIZONTAL_GENERAL;
import static io.github.hee9841.excel.style.align.alignment.VerticalAlignment.VERTICAL_CENTER;
import static org.mockito.BDDMockito.then;

import io.github.hee9841.excel.style.align.CustomExcelAlign;
import io.github.hee9841.excel.style.align.DefaultExcelAlign;
import io.github.hee9841.excel.style.align.alignment.HorizontalAlignment;
import io.github.hee9841.excel.style.align.alignment.VerticalAlignment;
import io.github.hee9841.excel.style.border.BorderStyle;
import io.github.hee9841.excel.style.border.DefaultExcelBorder;
import io.github.hee9841.excel.style.color.IndexedColors;
import io.github.hee9841.excel.style.color.IndexedExcelColor;
import io.github.hee9841.excel.style.color.RgbExcelColor;
import io.github.hee9841.excel.style.configurer.ExcelCellStyleConfigurer;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;


@ExtendWith(MockitoExtension.class)
class CustomExcelCellStyleImplementationTest {

    @Mock
    CellStyle cellStyle;

    @DisplayName("커스텀 셀 스타일 : IndexedColors, DefaultAlign, AllBorde 적용 테스트")
    @Test
    void IndexColor_DefaultAlign_DefaultBorder_Test() {
        //given
        IndexedColors grey25PercentColor = IndexedColors.GREY_25_PERCENT;
        DefaultExcelAlign align = DefaultExcelAlign.GENERAL_CENTER;
        BorderStyle borderThin = BorderStyle.THIN;

        class IndexedColorDefaultAlignAllBorder extends CustomExcelCellStyle {

            @Override
            public void configure(ExcelCellStyleConfigurer configurer) {
                configurer.excelColor(IndexedExcelColor.of(grey25PercentColor));
                configurer.excelAlign(align);
                configurer.excelBorder(DefaultExcelBorder.all(BorderStyle.THIN));
            }
        }

        //when
        new IndexedColorDefaultAlignAllBorder().apply(cellStyle);

        //then
        //color
        then(cellStyle).should().setFillForegroundColor(grey25PercentColor.getIndex());
        then(cellStyle).should().setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //align
        then(cellStyle).should()
            .setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL);
        then(cellStyle).should()
            .setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

        //border
        then(cellStyle).should().setBorderTop(borderThin.getBorderStyle());
        then(cellStyle).should().setBorderBottom(borderThin.getBorderStyle());
        then(cellStyle).should().setBorderLeft(borderThin.getBorderStyle());
        then(cellStyle).should().setBorderRight(borderThin.getBorderStyle());

        then(cellStyle).shouldHaveNoMoreInteractions();

    }

    @DisplayName("커스텀 셀 스타일 : RGB, CustomAlign, BuilderBorder 적용 테스트")
    @Test
    void RGB_CustomAlign_BuilderBorder_Test() {
        //given
        int red = 100, green = 100, blue = 100;
        HorizontalAlignment horizontal = HORIZONTAL_GENERAL;
        VerticalAlignment vertical = VERTICAL_CENTER;
        BorderStyle borderThin = BorderStyle.THIN;

        class RgbColorCustomAlignBuilderBorder extends CustomExcelCellStyle {

            @Override
            public void configure(ExcelCellStyleConfigurer configurer) {
                configurer.excelColor(RgbExcelColor.rgb(red, green, blue));
                configurer.excelAlign(CustomExcelAlign.of(horizontal, vertical));
                configurer.excelBorder(DefaultExcelBorder.builder().top(borderThin).build());
            }

        }

        //when
        new RgbColorCustomAlignBuilderBorder().apply(cellStyle);

        //then
        //color
        then(cellStyle).should().setFillForegroundColor(
            new XSSFColor(new byte[]{(byte) red, (byte) green, (byte) blue}));
        then(cellStyle).should().setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //align
        then(cellStyle).should().setAlignment(horizontal.getAlign());
        then(cellStyle).should().setVerticalAlignment(vertical.getAlign());

        //border
        then(cellStyle).should().setBorderTop(borderThin.getBorderStyle());

        then(cellStyle).shouldHaveNoMoreInteractions();
    }

}
