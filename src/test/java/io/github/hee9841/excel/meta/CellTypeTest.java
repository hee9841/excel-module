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
class CellTypeTest {


    @Test
    @DisplayName("from 메서드는 필드 타입에 맞는 CellType을 반환해야 한다")
    void from_ShouldReturnMatchingCellType() {
        //NUMBER
        assertEquals(CellType.NUMBER, CellType.from(Integer.class));
        assertEquals(CellType.NUMBER, CellType.from(int.class));
        assertEquals(CellType.NUMBER, CellType.from(Double.class));
        assertEquals(CellType.NUMBER, CellType.from(double.class));
        assertEquals(CellType.NUMBER, CellType.from(Float.class));
        assertEquals(CellType.NUMBER, CellType.from(float.class));
        assertEquals(CellType.NUMBER, CellType.from(Long.class));
        assertEquals(CellType.NUMBER, CellType.from(long.class));
        assertEquals(CellType.NUMBER, CellType.from(Short.class));
        assertEquals(CellType.NUMBER, CellType.from(short.class));
        assertEquals(CellType.NUMBER, CellType.from(Byte.class));
        assertEquals(CellType.NUMBER, CellType.from(byte.class));

        //BOOLEAN
        assertEquals(CellType.BOOLEAN, CellType.from(Boolean.class));
        assertEquals(CellType.BOOLEAN, CellType.from(boolean.class));

        //STRING
        assertEquals(CellType.STRING, CellType.from(String.class));
        assertEquals(CellType.STRING, CellType.from(Character.class));
        assertEquals(CellType.STRING, CellType.from(char.class));

        //Enum
        assertEquals(CellType.ENUM, CellType.from(MyEnum.class));

        //날짜 관련
        assertEquals(CellType.DATE, CellType.from(Date.class));
        assertEquals(CellType.LOCAL_DATE, CellType.from(LocalDate.class));
        assertEquals(CellType.LOCAL_DATE_TIME, CellType.from(LocalDateTime.class));
    }

    @Test
    @DisplayName("지원하지 않는 타입에 대해 from 메서드는 _NONE을 반환해야 한다")
    void from_ShouldReturnNoneForUnsupportedType() {
        // given & when & then
        assertEquals(CellType._NONE, CellType.from(Object.class));
    }

    @Test
    @DisplayName("CellType의 허용된 타입에 필드 타입이 있을경우, 해당 CellType을 반환해야한다.")
    void shouldReturnTargetCellTypeWhenMatched() {
        assertEquals(CellType.NUMBER,
            CellType.findMatchingCellType(Integer.class, CellType.NUMBER));
        assertEquals(CellType.STRING,
            CellType.findMatchingCellType(String.class, CellType.STRING));
        assertEquals(CellType.BOOLEAN,
            CellType.findMatchingCellType(Boolean.class, CellType.BOOLEAN));
        //formal 타입
        assertEquals(CellType.FORMULA,
            CellType.findMatchingCellType(String.class, CellType.FORMULA));
    }

    @Test
    @DisplayName("필드 타입과 CellType이 일치하지 않으면 _NONE을 반환해야 한다")
    void shouldReturnNoneWhenNotMatched() {
        // given & when & then
        assertEquals(CellType._NONE,
            CellType.findMatchingCellType(String.class, CellType.NUMBER));
        assertEquals(CellType._NONE,
            CellType.findMatchingCellType(Integer.class, CellType.STRING));
    }

    @DisplayName("CellType에 맞게 cell 값을 set해야한다.")
    @ParameterizedTest
    @MethodSource("cellValueData")
    void shouldSetCellValueByCellType(CellType cellType, Object value) {
        Cell cell = Mockito.mock(Cell.class);

        //when
        cellType.setCellValueByCellType(cell, value);

        //then
        switch (cellType) {
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
                CellType.NUMBER,
                1
            ),
            Arguments.of(
                CellType.NUMBER,
                1.0f
            ),
            Arguments.of(
                CellType.NUMBER,
                1L
            ),
            Arguments.of(
                CellType.BOOLEAN,
                true
            ),
            Arguments.of(
                CellType.STRING,
                "String"
            ),
            Arguments.of(
                CellType.STRING,
                'C'
            ),
            Arguments.of(
                CellType.FORMULA,
                "SUM(A1:A7)"
            ),
            Arguments.of(
                CellType.ENUM,
                MyEnum.FEMALE
            ),
            Arguments.of(
                CellType.DATE,
                new Date()
            ),
            Arguments.of(
                CellType.LOCAL_DATE,
                LocalDate.now()
            ),
            Arguments.of(
                CellType.LOCAL_DATE_TIME,
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
