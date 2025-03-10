package io.github.hee9841.excel.style.align;

import static org.mockito.BDDMockito.then;

import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

@ExtendWith(MockitoExtension.class)
class DefaultExcelAlignTest {


    @Mock
    private CellStyle cellStyle;


    @ParameterizedTest
    @MethodSource("generateAlignData")
    void alignApplyShouldSetCorrectAlignment(
        DefaultExcelAlign align,
        HorizontalAlignment poiHorizontal,
        VerticalAlignment poiVertical
    ) {

        // when
        align.applyAlign(cellStyle);

        // then
        then(cellStyle).should().setAlignment(poiHorizontal);
        then(cellStyle).should().setVerticalAlignment(poiVertical);
    }

    static Stream<Arguments> generateAlignData() {
        return Stream.of(
            Arguments.of(
                DefaultExcelAlign.GENERAL_CENTER,
                HorizontalAlignment.GENERAL,
                VerticalAlignment.CENTER
            ),
            Arguments.of(
                DefaultExcelAlign.LEFT_CENTER,
                HorizontalAlignment.LEFT,
                VerticalAlignment.CENTER
            ),
            Arguments.of(
                DefaultExcelAlign.CENTER_CENTER,
                HorizontalAlignment.CENTER,
                VerticalAlignment.CENTER
            ),
            Arguments.of(
                DefaultExcelAlign.RIGHT_CENTER,
                HorizontalAlignment.RIGHT,
                VerticalAlignment.CENTER
            ),
            Arguments.of(
                DefaultExcelAlign.FILL_CENTER,
                HorizontalAlignment.FILL,
                VerticalAlignment.CENTER
            ),
            Arguments.of(
                DefaultExcelAlign.JUSTIFY_CENTER,
                HorizontalAlignment.JUSTIFY,
                VerticalAlignment.CENTER
            ),
            Arguments.of(
                DefaultExcelAlign.CENTER_SELECTION_CENTER,
                HorizontalAlignment.CENTER_SELECTION,
                VerticalAlignment.CENTER
            ),
            Arguments.of(
                DefaultExcelAlign.DISTRIBUTED_CENTER,
                HorizontalAlignment.DISTRIBUTED,
                VerticalAlignment.CENTER
            )
        );
    }

}
