package io.github.hee9841.excel.style.border;

import static org.mockito.BDDMockito.then;

import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

@ExtendWith(MockitoExtension.class)
class DefaultExcelBorderTest {

    @Mock
    private CellStyle cellStyle;

    @DisplayName("all 메서드는 모든 테두리를 적용한다.")
    @Test
    void allMethod_ShouldApplyAllBorder() {
        //given
        DefaultExcelBorder border = DefaultExcelBorder.all(BorderStyle.THIN);

        //when
        border.applyAllBorder(cellStyle);

        //then
        then(cellStyle).should().setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        then(cellStyle).should().setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        then(cellStyle).should().setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        then(cellStyle).should().setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
    }

    @DisplayName("builder 패턴 적용 시, 테두리가 적용되는 경우를 테스트한다.")
    @ParameterizedTest
    @MethodSource("generateBuilderData")
    void builderMethod_ShouldApplyBorder(
        DefaultExcelBorder border,
        org.apache.poi.ss.usermodel.BorderStyle top,
        org.apache.poi.ss.usermodel.BorderStyle bottom,
        org.apache.poi.ss.usermodel.BorderStyle left,
        org.apache.poi.ss.usermodel.BorderStyle right
    ) {

        //when
        border.applyAllBorder(cellStyle);

        //then
        if (top != null) {
            then(cellStyle).should().setBorderTop(top);
        }
        if (bottom != null) {
            then(cellStyle).should().setBorderBottom(bottom);
        }
        if (left != null) {
            then(cellStyle).should().setBorderLeft(left);
        }
        if (right != null) {
            then(cellStyle).should().setBorderRight(right);
        }

        then(cellStyle).shouldHaveNoMoreInteractions();
    }

    static Stream<Arguments> generateBuilderData() {
        return Stream.of(
            Arguments.of(
                DefaultExcelBorder.builder()
                    .top(BorderStyle.THIN)
                    .bottom(BorderStyle.MEDIUM)
                    .left(BorderStyle.DASHED)
                    .right(BorderStyle.DOTTED).build(),
                org.apache.poi.ss.usermodel.BorderStyle.THIN,
                org.apache.poi.ss.usermodel.BorderStyle.MEDIUM,
                org.apache.poi.ss.usermodel.BorderStyle.DASHED,
                org.apache.poi.ss.usermodel.BorderStyle.DOTTED
            ),
            Arguments.of(
                DefaultExcelBorder.builder()
                    .bottom(BorderStyle.MEDIUM)
                    .left(BorderStyle.DASHED)
                    .right(BorderStyle.DOTTED).build(),
                null,
                org.apache.poi.ss.usermodel.BorderStyle.MEDIUM,
                org.apache.poi.ss.usermodel.BorderStyle.DASHED,
                org.apache.poi.ss.usermodel.BorderStyle.DOTTED
            ),
            Arguments.of(
                DefaultExcelBorder.builder()
                    .top(BorderStyle.THIN)
                    .left(BorderStyle.DASHED)
                    .right(BorderStyle.DOTTED).build(),
                org.apache.poi.ss.usermodel.BorderStyle.THIN,
                null,
                org.apache.poi.ss.usermodel.BorderStyle.DASHED,
                org.apache.poi.ss.usermodel.BorderStyle.DOTTED
            ),
            Arguments.of(
                DefaultExcelBorder.builder()
                    .top(BorderStyle.THIN)
                    .bottom(BorderStyle.MEDIUM)
                    .right(BorderStyle.DOTTED).build(),
                org.apache.poi.ss.usermodel.BorderStyle.THIN,
                org.apache.poi.ss.usermodel.BorderStyle.MEDIUM,
                null,
                org.apache.poi.ss.usermodel.BorderStyle.DOTTED
            ),
            Arguments.of(
                DefaultExcelBorder.builder()
                    .top(BorderStyle.THIN)
                    .bottom(BorderStyle.MEDIUM)
                    .left(BorderStyle.DASHED).build(),
                org.apache.poi.ss.usermodel.BorderStyle.THIN,
                org.apache.poi.ss.usermodel.BorderStyle.MEDIUM,
                org.apache.poi.ss.usermodel.BorderStyle.DASHED,
                null
            ),
            Arguments.of(
                DefaultExcelBorder.builder().build(),
                null,
                null,
                null,
                null
            )
        );
    }
}

