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
 * <p><b>Resource Management:</b> This interface extends {@link AutoCloseable} to ensure proper
 * resource cleanup. It is recommended to use try-with-resources when working with exporters,
 * especially since implementations like {@code SXSSFExporter} create temporary files:</p>
 * <pre>{@code
 * try (ExcelExporter<MyData> exporter = SXSSFExporter.builder(MyData.class, data).build()) {
 *     exporter.addRows(moreData);
 *     exporter.write(outputStream);
 * }
 * }</pre>
 *
 * <p><b>Thread Safety:</b> Implementations of this interface are NOT thread-safe.
 * A single instance should not be shared across multiple threads.
 * Each thread should create its own exporter instance.</p>
 *
 * @param <T> The type of data to be handled in the Excel file
 */
public interface ExcelExporter<T> extends AutoCloseable {

    /**
     * Writes the Excel file content to the specified output stream.
     * After this method completes, the exporter is closed and cannot be reused.
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

    /**
     * Closes the exporter and releases any resources associated with it.
     * This method is idempotent - calling it multiple times has no additional effect.
     */
    @Override
    void close();
}
