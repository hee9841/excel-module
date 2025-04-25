package io.github.hee9841.excel.style.align;

import static org.mockito.BDDMockito.then;

import io.github.hee9841.excel.style.align.alignment.ExcelHorizontalAlignment;
import io.github.hee9841.excel.style.align.alignment.ExcelVerticalAlignment;
import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

@ExtendWith(MockitoExtension.class)
class CustomExcelAlignTest {

    @Mock
    private CellStyle cellStyle;


    @DisplayName("from(HorizontalAlignment) 메서드는 horizontal만 설정되어야 한다.")
    @ParameterizedTest
    @MethodSource("generateHorizontalData")
    void from_HorizontalAlignment(
        ExcelHorizontalAlignment horizontal,
        org.apache.poi.ss.usermodel.HorizontalAlignment poiHorizontal
    ) {

        CustomExcelAlign align = CustomExcelAlign.from(horizontal);

        align.applyAlign(cellStyle);

        then(cellStyle).should().setAlignment(poiHorizontal);
        then(cellStyle).should()
            .setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
    }

    @DisplayName("from(VerticalAlignment) 메서드는 vertical만 설정되어야 한다.")
    @ParameterizedTest
    @MethodSource("generateVerticalData")
    void from_VerticalAlignment(
        ExcelVerticalAlignment vertical,
        org.apache.poi.ss.usermodel.VerticalAlignment poiVertical
    ) {

        CustomExcelAlign align = CustomExcelAlign.from(vertical);

        align.applyAlign(cellStyle);

        then(cellStyle).should().setVerticalAlignment(poiVertical);
        then(cellStyle).shouldHaveNoMoreInteractions();
    }

    @DisplayName("of 메서드로 생성 시, horizontal와 vertical 모두 설정되어야 한다.")
    @ParameterizedTest
    @MethodSource("generateFoAlignData")
    void of_ShouldSetBothHorizontalAndVerticalAlignments(
        ExcelHorizontalAlignment horizontal,
        ExcelVerticalAlignment vertical,
        org.apache.poi.ss.usermodel.HorizontalAlignment poiHorizontal,
        org.apache.poi.ss.usermodel.VerticalAlignment poiVertical
    ) {

        CustomExcelAlign align = CustomExcelAlign.of(horizontal, vertical);

        align.applyAlign(cellStyle);

        then(cellStyle).should().setAlignment(poiHorizontal);
        then(cellStyle).should().setVerticalAlignment(poiVertical);
    }

    static Stream<Arguments> generateHorizontalData() {
        return Stream.of(
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_GENERAL,
                org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_LEFT,
                org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_CENTER,
                org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_RIGHT,
                org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_FILL,
                org.apache.poi.ss.usermodel.HorizontalAlignment.FILL
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_JUSTIFY,
                org.apache.poi.ss.usermodel.HorizontalAlignment.JUSTIFY
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_CENTER_SELECTION,
                org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER_SELECTION
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_DISTRIBUTED,
                org.apache.poi.ss.usermodel.HorizontalAlignment.DISTRIBUTED
            )
        );
    }

    static Stream<Arguments> generateVerticalData() {
        return Stream.of(
            Arguments.of(
                ExcelVerticalAlignment.VERTICAL_TOP,
                org.apache.poi.ss.usermodel.VerticalAlignment.TOP
            ),
            Arguments.of(
                ExcelVerticalAlignment.VERTICAL_CENTER,
                org.apache.poi.ss.usermodel.VerticalAlignment.CENTER
            ),
            Arguments.of(
                ExcelVerticalAlignment.VERTICAL_BOTTOM,
                org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM
            ),
            Arguments.of(
                ExcelVerticalAlignment.VERTICAL_JUSTIFY,
                org.apache.poi.ss.usermodel.VerticalAlignment.JUSTIFY
            ),
            Arguments.of(
                ExcelVerticalAlignment.VERTICAL_DISTRIBUTED,
                org.apache.poi.ss.usermodel.VerticalAlignment.DISTRIBUTED
            )
        );
    }

    static Stream<Arguments> generateFoAlignData() {
        return Stream.of(
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_GENERAL,
                ExcelVerticalAlignment.VERTICAL_TOP,
                org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL,
                org.apache.poi.ss.usermodel.VerticalAlignment.TOP
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_GENERAL,
                ExcelVerticalAlignment.VERTICAL_CENTER,
                org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL,
                org.apache.poi.ss.usermodel.VerticalAlignment.CENTER
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_LEFT,
                ExcelVerticalAlignment.VERTICAL_BOTTOM,
                org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT,
                org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_LEFT,
                ExcelVerticalAlignment.VERTICAL_JUSTIFY,
                org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT,
                org.apache.poi.ss.usermodel.VerticalAlignment.JUSTIFY
            ),
            Arguments.of(
                ExcelHorizontalAlignment.HORIZONTAL_LEFT,
                ExcelVerticalAlignment.VERTICAL_DISTRIBUTED,
                org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT,
                org.apache.poi.ss.usermodel.VerticalAlignment.DISTRIBUTED
            )
        );
    }

}
