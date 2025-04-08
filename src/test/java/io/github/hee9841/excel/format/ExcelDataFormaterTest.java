package io.github.hee9841.excel.format;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.BDDMockito.given;
import static org.mockito.BDDMockito.then;
import static org.mockito.Mockito.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;


class ExcelDataFormaterTest {

    @Test
    @DisplayName("of 메서드는 DataFormat과 패턴으로 ExcelDataFormater 인스턴스를 생성한다")
    void create_instance() {
        // given
        DataFormat dataFormat = mock(DataFormat.class);
        String pattern = "yyyy-MM-dd";

        // when
        ExcelDataFormater formater = ExcelDataFormater.of(dataFormat, pattern);

        // then
        assertNotNull(formater);
    }


    @Nested
    @DisplayName("apply 메서드는")
    @ExtendWith(MockitoExtension.class)
    class Apply {

        @Mock
        private CellStyle cellStyle;

        @Mock
        private DataFormat dataFormat;


        @Test
        @DisplayName("패턴이 있을 때, 셀 스타일에 데이터 포맷을 적용한다")
        void apply_format_when_pattern_is_not_none() {
            // given
            String pattern = "yyyy-MM-dd";
            short patternShortValue = (short) 1;
            given(dataFormat.getFormat(pattern)).willReturn(patternShortValue);

            ExcelDataFormater formater = ExcelDataFormater.of(dataFormat, pattern);

            // when
            formater.apply(cellStyle);

            // then
            then(cellStyle).should().setDataFormat(patternShortValue);
        }

        @Test
        @DisplayName("패턴이 none일 때 셀 스타일에 데이터 포맷을 적용하지 않는다")
        void not_apply_format_when_pattern_is_none() {
            // given
            String pattern = CellFormats._NONE;
            ExcelDataFormater formater = ExcelDataFormater.of(dataFormat, pattern);

            // when
            formater.apply(cellStyle);

            // then
            then(cellStyle).should(never()).setDataFormat(anyShort());
        }

        @Test
        @DisplayName("패턴이 null일 때 셀 스타일에 데이터 포맷을 적용하지 않는다")
        void not_apply_format_when_pattern_is_null() {
            // given
            String pattern = null;

            ExcelDataFormater formater = ExcelDataFormater.of(dataFormat, pattern);

            // when
            formater.apply(cellStyle);

            // then
            then(cellStyle).should(never()).setDataFormat(anyShort());
        }

        @Test
        @DisplayName("패턴이 빈 문자열일 때 셀 스타일에 데이터 포맷을 적용하지 않는다")
        void not_apply_format_when_pattern_is_empty() {
            // given
            String pattern = " ";
            ExcelDataFormater formater = ExcelDataFormater.of(dataFormat, pattern);

            // when
            formater.apply(cellStyle);

            // then
            then(cellStyle).should(never()).setDataFormat(anyShort());
        }
    }
}
