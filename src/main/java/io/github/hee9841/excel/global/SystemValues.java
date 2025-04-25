package io.github.hee9841.excel.global;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.Set;
import java.util.stream.Collectors;

public class SystemValues {

    public static final Set<Class<?>> ALLOWED_FIELD_TYPES = Collections.unmodifiableSet(
        new HashSet<>(Arrays.asList(
            String.class,
            Character.class, char.class,

            // numeric
            Byte.class, byte.class,
            Short.class, short.class,
            Integer.class, int.class,
            Long.class, long.class,
            Float.class, float.class,
            Double.class, double.class,

            // boolean
            Boolean.class, boolean.class,

            // date/time type
            LocalDate.class,
            LocalDateTime.class,
            Date.class,
            java.sql.Date.class,

            Enum.class
        ))
    );

    public static final String ALLOWED_FIELD_TYPES_STRING = ALLOWED_FIELD_TYPES.stream()
        .map(Class::getSimpleName)
        .collect(Collectors.joining(", "));

    private SystemValues() {
    }
}
