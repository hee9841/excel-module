package io.github.hee9841.excel.core;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public interface ExcelFile<T> {

    void write(OutputStream stream) throws IOException;

    void addRows(List<T> data);
}
