package io.github.hee9841.excel.format;

public class CellFormats {

    public static final String _NONE = "General";

    public static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";
    public static final String DEFAULT_DATE_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";
    public static final String DEFAULT_LOCAL_TIME_FORMAT = "HH:mm:ss";

    public static final String THOUSAND_SEPARATED_NUMBER_FORMAT = "#,##0";
    public static final String THOUSAND_SEPARATED_DECIMAL_FORMAT = "#,##0.00";

    public static final String DOLLAR_FORMAT = "$#,##0.00";
    public static final String KR_WON_FORMAT = "₩#,##0";


    private CellFormats() {
    }

    public static boolean isNone(String format) {
        return format == null || format.equals(_NONE) || format.trim().isEmpty();
    }
}
