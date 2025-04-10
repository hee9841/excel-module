package io.github.hee9841.excel.meta;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.BDDMockito.then;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.Cell;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
import org.mockito.Mockito;
import org.mockito.junit.jupiter.MockitoExtension;


@ExtendWith(MockitoExtension.class)
class ColumnDataTypeTest {


    @Test
    @DisplayName("from 메서드는 필드 타입에 맞는 CellType을 반환해야 한다")
    void from_ShouldReturnMatchingCellType() {
        //NUMBER
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(Integer.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(int.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(Double.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(double.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(Float.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(float.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(Long.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(long.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(Short.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(short.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(Byte.class));
        assertEquals(ColumnDataType.NUMBER, ColumnDataType.from(byte.class));

        //BOOLEAN
        assertEquals(ColumnDataType.BOOLEAN, ColumnDataType.from(Boolean.class));
        assertEquals(ColumnDataType.BOOLEAN, ColumnDataType.from(boolean.class));

        //STRING
        assertEquals(ColumnDataType.STRING, ColumnDataType.from(String.class));
        assertEquals(ColumnDataType.STRING, ColumnDataType.from(Character.class));
        assertEquals(ColumnDataType.STRING, ColumnDataType.from(char.class));

        //Enum
        assertEquals(ColumnDataType.ENUM, ColumnDataType.from(MyEnum.class));

        //날짜 관련
        assertEquals(ColumnDataType.DATE, ColumnDataType.from(Date.class));
        assertEquals(ColumnDataType.LOCAL_DATE, ColumnDataType.from(LocalDate.class));
        assertEquals(ColumnDataType.LOCAL_DATE_TIME, ColumnDataType.from(LocalDateTime.class));
    }

    @Test
    @DisplayName("지원하지 않는 타입에 대해 from 메서드는 _NONE을 반환해야 한다")
    void from_ShouldReturnNoneForUnsupportedType() {
        // given & when & then
        assertEquals(ColumnDataType._NONE, ColumnDataType.from(Object.class));
    }

    @Test
    @DisplayName("CellType의 허용된 타입에 필드 타입이 있을경우, 해당 CellType을 반환해야한다.")
    void shouldReturnTargetCellTypeWhenMatched() {
        assertEquals(ColumnDataType.NUMBER,
            ColumnDataType.findMatchingCellType(Integer.class, ColumnDataType.NUMBER));
        assertEquals(ColumnDataType.STRING,
            ColumnDataType.findMatchingCellType(String.class, ColumnDataType.STRING));
        assertEquals(ColumnDataType.BOOLEAN,
            ColumnDataType.findMatchingCellType(Boolean.class, ColumnDataType.BOOLEAN));
        //formal 타입
        assertEquals(ColumnDataType.FORMULA,
            ColumnDataType.findMatchingCellType(String.class, ColumnDataType.FORMULA));
    }

    @Test
    @DisplayName("필드 타입과 CellType이 일치하지 않으면 _NONE을 반환해야 한다")
    void shouldReturnNoneWhenNotMatched() {
        // given & when & then
        assertEquals(ColumnDataType._NONE,
            ColumnDataType.findMatchingCellType(String.class, ColumnDataType.NUMBER));
        assertEquals(ColumnDataType._NONE,
            ColumnDataType.findMatchingCellType(Integer.class, ColumnDataType.STRING));
    }

    @DisplayName("CellType에 맞게 cell 값을 set해야한다.")
    @ParameterizedTest
    @MethodSource("cellValueData")
    void shouldSetCellValueByCellType(ColumnDataType columnDataType, Object value) {
        Cell cell = Mockito.mock(Cell.class);

        //when
        columnDataType.setCellValueByCellType(cell, value);

        //then
        switch (columnDataType) {
            case NUMBER:
                then(cell).should().setCellValue(Double.parseDouble(String.valueOf(value)));
                break;
            case BOOLEAN:
                then(cell).should().setCellValue((boolean) value);
                break;
            case ENUM:
                then(cell).should().setCellValue(value != null ? value.toString() : "");
                break;
            case FORMULA:
                then(cell).should().setCellFormula(String.valueOf(value));
                break;
            case DATE:
                then(cell).should().setCellValue((Date) value);
                break;
            case LOCAL_DATE:
                then(cell).should().setCellValue((LocalDate) value);
                break;
            case LOCAL_DATE_TIME:
                then(cell).should().setCellValue((LocalDateTime) value);
                break;
            default:
                then(cell).should().setCellValue(String.valueOf(value));
                break;
        }

    }


    static Stream<Arguments> cellValueData() {
        return Stream.of(
            Arguments.of(
                ColumnDataType.NUMBER,
                1
            ),
            Arguments.of(
                ColumnDataType.NUMBER,
                1.0f
            ),
            Arguments.of(
                ColumnDataType.NUMBER,
                1L
            ),
            Arguments.of(
                ColumnDataType.BOOLEAN,
                true
            ),
            Arguments.of(
                ColumnDataType.STRING,
                "String"
            ),
            Arguments.of(
                ColumnDataType.STRING,
                'C'
            ),
            Arguments.of(
                ColumnDataType.FORMULA,
                "SUM(A1:A7)"
            ),
            Arguments.of(
                ColumnDataType.ENUM,
                MyEnum.FEMALE
            ),
            Arguments.of(
                ColumnDataType.DATE,
                new Date()
            ),
            Arguments.of(
                ColumnDataType.LOCAL_DATE,
                LocalDate.now()
            ),
            Arguments.of(
                ColumnDataType.LOCAL_DATE_TIME,
                LocalDateTime.now()
            )
        );
    }


    enum MyEnum {
        FEMALE("F"),
        MALE("M"),
        ;
        private final String gender;

        MyEnum(String gender) {
            this.gender = gender;
        }


        @Override
        public String toString() {
            return this.gender;
        }
    }


}
