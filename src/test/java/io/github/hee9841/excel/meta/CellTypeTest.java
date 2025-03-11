package io.github.hee9841.excel.meta;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.Nested;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

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



    @Nested
    @DisplayName("findMatchingCellType 메서드 테스트")
    class FindMatchingCellTypeTest {
        
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
    }

}
