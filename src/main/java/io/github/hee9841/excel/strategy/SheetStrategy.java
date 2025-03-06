package io.github.hee9841.excel.strategy;

import static org.apache.commons.compress.archivers.zip.Zip64Mode.Always;
import static org.apache.commons.compress.archivers.zip.Zip64Mode.AsNeeded;

import java.util.Objects;
import org.apache.commons.compress.archivers.zip.Zip64Mode;

public enum SheetStrategy {
    ONE_SHEET(AsNeeded),
    MULTI_SHEET(Always),
    AUTO(null);


    private final Zip64Mode zip64Mode;

    SheetStrategy(Zip64Mode zip64Mode) {
        this.zip64Mode = zip64Mode;
    }

    public static boolean isOneSheet(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, ONE_SHEET);
    }

    public static boolean isAuto(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, AUTO);
    }

    public static boolean isMultiSheet(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, MULTI_SHEET);
    }

    public Zip64Mode getZip64Mode() {
        return zip64Mode;
    }
}
