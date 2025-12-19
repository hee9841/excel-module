package io.github.hee9841.excel.style;

import io.github.hee9841.excel.style.configurer.ExcelCellStyleConfigurer;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Abstract base class implementing the ExcelCellStyle interface using the Template Method pattern.
 *
 * <p>This class provides a structured approach to define cell styles by delegating
 * the configuration details to subclasses while handling the common application logic.</p>
 *
 * <p>Subclasses need to implement the {@link #configure(ExcelCellStyleConfigurer)} method
 * to define specific style configurations (colors, borders, alignment, etc.).</p>
 *
 * @see io.github.hee9841.excel.style.ExcelCellStyle
 * @see io.github.hee9841.excel.style.configurer.ExcelCellStyleConfigurer
 */
public abstract class CustomExcelCellStyle implements ExcelCellStyle {

    private ExcelCellStyleConfigurer configurer;

    /**
     * Template method to be implemented by subclasses to define specific style configurations.
     * This method is invoked lazily when the style is first applied to build the configurer.
     *
     * @param configurer the style configurer to be set up with specific style settings
     */
    public abstract void configure(ExcelCellStyleConfigurer configurer);

    /**
     * Applies all configured styles to the provided cell style.
     * This implementation delegates to the configurer to apply the configured styles.
     *
     * @param cellStyle the cell style to which styles will be applied
     */
    @Override
    public void apply(CellStyle cellStyle) {
        getConfigurer().configure(cellStyle);
    }

    private ExcelCellStyleConfigurer getConfigurer() {
        if (configurer == null) {
            ExcelCellStyleConfigurer initialized = new ExcelCellStyleConfigurer();
            configure(initialized);
            configurer = initialized;
        }
        return configurer;
    }

}
