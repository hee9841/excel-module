package io.github.hee9841.excel.global;


import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class AllowedTypes {
    public static final List<Class<?>> NUMBER_TYPES = Collections.unmodifiableList(Arrays.asList(
        Byte.class, byte.class,
        Short.class, short.class,
        Integer.class, int.class,
        Long.class, long.class,
        Float.class, float.class,
        Double.class, double.class
    ));

    public static final List<Class<?>> STRING_TYPES = Collections.unmodifiableList(Arrays.asList(
        String.class,
        Character.class, char.class
    ));

    public static final List<Class<?>> BOOLEAN_TYPES = Collections.unmodifiableList(Arrays.asList(
        Boolean.class, boolean.class
    ));

    public static final List<Class<?>> ENUM_TYPES = Collections.singletonList(Enum.class);
    public static final List<Class<?>> FORMULA_TYPES = Collections.singletonList(String.class);

    public static final List<Class<?>> LOCAL_DATE_TYPES = Collections.singletonList(LocalDate.class);
    public static final List<Class<?>> LOCAL_DATE_TIME_TYPES = Collections.singletonList(LocalDateTime.class);

    public static final List<Class<?>> DATE_TYPES = Collections.unmodifiableList(Arrays.asList(
        Date.class, java.sql.Date.class
    ));

    public static final Set<Class<?>> ALLOWED_FIELD_TYPES = Collections.unmodifiableSet(
        Stream.of(
            NUMBER_TYPES,
            STRING_TYPES,
            BOOLEAN_TYPES,
            ENUM_TYPES,
            FORMULA_TYPES,
            LOCAL_DATE_TYPES,
            LOCAL_DATE_TIME_TYPES,
            DATE_TYPES
        ).flatMap(List::stream).collect(Collectors.toSet())
    );

    public static final String ALLOWED_FIELD_TYPES_STRING = ALLOWED_FIELD_TYPES.stream()
        .map(Class::getSimpleName)
        .collect(Collectors.joining(", "));


    private AllowedTypes() {
    }
}
