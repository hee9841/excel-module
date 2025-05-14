package io.github.hee9841.excel.core.exporter;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * Core interface for Excel file operations in the library.
 * This interface defines the basic contract for Excel file handling operations.
 *
 * <p>Implementations of this interface provide functionality for:</p>
 * <ul>
 *     <li>Writing Excel data to an output stream</li>
 *     <li>Adding rows of data to the Excel file</li>
 * </ul>
 *
 * @param <T> The type of data to be handled in the Excel file
 */
public interface ExcelExporter<T> {

    /**
     * Writes the Excel file content to the specified output stream.
     *
     * @param stream The output stream to write the Excel file to
     * @throws IOException if an I/O error occurs during writing
     */
    void write(OutputStream stream) throws IOException;

    /**
     * Adds a list of data rows to the Excel file.
     *
     * @param data The list of data objects to be added as rows
     */
    void addRows(List<T> data);
}
