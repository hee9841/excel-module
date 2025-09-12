package io.github.hee9841.excel.strategy;

import static org.apache.commons.compress.archivers.zip.Zip64Mode.Always;
import static org.apache.commons.compress.archivers.zip.Zip64Mode.AsNeeded;


import java.util.Objects;
import org.apache.commons.compress.archivers.zip.Zip64Mode;

/**
 * Strategy enum that determines how Excel sheets are created and managed during export operations.
 * <p>
 * This enum controls whether data is exported to a single sheet or multiple sheets, which affects
 * the Zip64 mode used when creating large Excel files. Each strategy is associated with a specific
 * {@link Zip64Mode} that optimizes file handling based on the expected data volume.
 * <p>
 * The Zip64 mode affects how Apache POI handles large Excel files:
 * <ul>
 *   <li>{@code ONE_SHEET}: Uses {@code AsNeeded} mode, which only enables Zip64 extensions when required</li>
 *   <li>{@code MULTI_SHEET}: Uses {@code Always} mode, which always enables Zip64 extensions for better
 *       handling of multiple sheets</li>
 * </ul>
 *
 * @see org.apache.commons.compress.archivers.zip.Zip64Mode
 * @see io.github.hee9841.excel.annotation.Excel
 */
public enum SheetStrategy {
    /**
     * Strategy for exporting data to a single sheet.
     * Uses the {@code AsNeeded} Zip64 mode which only enables Zip64 extensions when necessary.
     * This is more efficient for smaller datasets that fit within a single sheet.
     */
    ONE_SHEET(AsNeeded),

    /**
     * Strategy for exporting data across multiple sheets.
     * Uses the {@code Always} Zip64 mode which always enables Zip64 extensions.
     * This is recommended for larger datasets that need to be split across multiple sheets.
     */
    MULTI_SHEET(Always),
    ;


    private final Zip64Mode zip64Mode;

    SheetStrategy(Zip64Mode zip64Mode) {
        this.zip64Mode = zip64Mode;
    }

    public static boolean isOneSheet(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, ONE_SHEET);
    }

    public static boolean isMultiSheet(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, MULTI_SHEET);
    }

    public Zip64Mode getZip64Mode() {
        return zip64Mode;
    }
}
