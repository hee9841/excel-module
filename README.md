# Excel Module

[![Maven Central](https://img.shields.io/maven-central/v/io.github.hee9841/excel-module.svg)](https://search.maven.org/artifact/io.github.hee9841/excel-module)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Build Status](https://github.com/hee9841/excel-module/workflows/Java%20CI%20with%20Gradle/badge.svg)](https://github.com/hee9841/excel-module/actions)
[![Coverage Status](https://coveralls.io/repos/github/hee9841/excel-module/badge.svg?branch=master)](https://coveralls.io/github/hee9841/excel-module?branch=master)

A Java library that helps you work with Excel files easily.

## Table of Contents
- [Excel Module](#excel-module)
  - [Table of Contents](#table-of-contents)
  - [Introduction](#introduction)
  - [Requirements](#requirements)
  - [Installation](#installation)
    - [Gradle](#gradle)
  - [Quick Start](#quick-start)
  - [Features \& Specifications](#features--specifications)
    - [Core Features](#core-features)
    - [Annotations](#annotations)
    - [Custom Cell Styling Options\*\*](#custom-cell-styling-options)
    - [Modes](#modes)
  - [Annotations](#annotations-1)
    - [@Excel (Class Level)](#excel-class-level)
    - [@ExcelColumn (Field Level)](#excelcolumn-field-level)
    - [@ExcelColumnStyle](#excelcolumnstyle)
  - [Custom Cell Style](#custom-cell-style)
    - [1. Using Enum classes](#1-using-enum-classes)
    - [2. Using Custom Classes](#2-using-custom-classes)
  - [Modes](#modes-1)
    - [CellTypeMode](#celltypemode)
    - [ColumnIndexMode](#columnindexmode)
    - [DataFormatMode](#dataformatmode)
    - [SheetMode](#sheetmode)
  - [Troubleshooting \& FAQ](#troubleshooting--faq)
    - [Common Issues](#common-issues)
  - [API Documentation](#api-documentation)
  - [License](#license)
  - [Contact](#contact)

## Introduction

This library is based on Apache POI and helps you to easily map Java DTO (Data Transfer Object) pattern classes to Excel files using annotations, simplifying the process of converting your data models to Excel format. It eliminates the need to write multiple lines of POI code by providing a simple, annotation-based approach.

**Key features include:**

- Excel file write operations
- Annotation-based Excel mapping
- Cell Styling and Data formatting
- Flexible extensibility through configurable modes

## Requirements

- Java 8 or higher (Java 23 supported)

## Installation

### Gradle

```gradle
dependencies {
    implementation 'io.github.hee9841.excel:excel-module:0.0.1'
    
    // Using annotationProcessor allows you to detect Excel annotation-related errors at compile time
    annotationProcessor 'io.github.hee9841.excel:excel-module:0.0.1'
}
```

## Quick Start

Here's a simple example to get you started:

```java
// 1. Define your data model with Excel annotations
@Excel(
    columnIndexMode = ColumnIndexMode.USER_DEFINED,
    cellTypeMode = CellTypeMode.AUTO,
    dataFormatMode = DataFormatMode.AUTO_BY_CELL_TYPE
)
public class Product {

    @ExcelColumn(headerName = "Product ID", columnIndex = 0)
    private Long id;

    @ExcelColumn(headerName = "Product Name", columnIndex = 1)
    private String name;

    @ExcelColumn(headerName = "Price", columnIndex = 2, format = "#,##0.00")
    private Double price;

    @ExcelColumn(headerName = "dateTime", 
        columnIndex = 3,
        columnColumnDataType = ColumnDataType.LOCAL_DATE_TIME
    )
    private LocalDateTime dateTime;

    // Constructors, getters, and setters
}

// 2. Create some data
List<Product> products = Arrays.asList(
    new Product(1L, "Laptop", 1299.99, LocalDateTime.now()),
    new Product(2L, "Smartphone", 899.99, LocalDateTime.now()),
    new Product(3L, "Headphones", 249.99, LocalDateTime.now())
);

// 3. Export to Excel

// Simple case: When you don't need to add more data after building
SXSSFExporter.builder(Product.class, products)
        .sheetMode(SheetMode.MULTI_SHEET) // Optional, MULTI_SHEET is default
        .maxRows(100)   // Optional, Max row of SpreadsheetVersion.EXCEL2007 is default
        .sheetName("Products") // Optional, if not specified sheets will be named Sheet0, Sheet1, etc.
        .build()
        .write(outputStream);

// Advanced case: When you need to add more data using addRows()
try (ExcelExporter<Product> exporter = SXSSFExporter.builder(Product.class, products)
         .sheetName("Products")
         .build()) {

    // Add more data if needed
    exporter.addRows(moreProducts);

    // Write to file or stream
    exporter.write(outputStream);
}  // Resources are automatically cleaned up
```

## Features & Specifications

This library provides several key features and specifications to help you work with Excel files:

### Core Features
- Excel file write operations with Apache POI
- Annotation-based Excel mapping for Java classes
- Customizable cell styling and data formatting
- Flexible modes for column indexing, cell types, and data formats

### Annotations
- `@Excel` - Class level configuration
- `@ExcelColumn` - Field level mapping
- `@ExcelColumnStyle` - Cell styling configuration

### Custom Cell Styling Options**
- Enum-based cell styles
- Custom style classes
- Pre-defined styles and formats

### Modes
- Column indexing modes
- Cell type determination modes
- Data format modes
- Sheet creation modes

See the sections below for detailed information about each feature.

## Annotations

This section details the annotations used to map your Java classes to Excel.

### @Excel (Class Level)

The `@Excel` annotation is applied at the class level and configures global Excel settings:

```java
@Excel(
    columnIndexMode = ColumnIndexMode.USER_DEFINED,
    cellTypeMode = CellTypeMode.AUTO,
    dataFormatMode = DataFormatMode.AUTO_BY_CELL_TYPE,
    headerStyle = @ExcelColumnStyle(cellStyleClass = MyHeaderEnumStyle.class, enumName="GRAY_25"),
    bodyStyle = @ExcelColumnStyle(cellStyleClass = MyBodyStyle.class)
)
public class UserDTO {
    // Fields with @ExcelColumn annotations
}
```

**Key parameters:**

| Parameter             | Description                             | Default                           |
|-----------------------|-----------------------------------------|-----------------------------------|
| `columnIndexMode` | Mode for determining column indices | `ColumnIndexMode.FIELD_ORDER` |
| `cellTypeMode`    | Mode for determining cell types     | `CellTypeMode.NONE`           |
| `dataFormatMode`  | Mode for applying data formats      | `DataFormatMode.NONE`         |
| `headerStyle`         | Default style for header cells          | *None*                            |
| `bodyStyle`           | Default style for body cells            | *None*                            |

> **Note:** 
> <br>- This annotation can only be applied to regular classes and records (not abstract classes, interfaces, or enums). 
> <br>- At least one field must have an `@ExcelColumn` annotation.

### @ExcelColumn (Field Level)

The `@ExcelColumn` annotation maps class fields to Excel columns:

```java
@ExcelColumn(
    headerName = "User ID",
    columnIndex = 0,
    columnColumnDataType = ColumnDataType.NUMBER,
    format = "#,##0",
    headerStyle = @ExcelColumnStyle(cellStyleClass = CustomHeaderStyle.class),
    bodyStyle = @ExcelColumnStyle(cellStyleClass = CustomBodyStyle.class)
)
private Long id;
```

**Key parameters:**

| Parameter              | Description                        | Required?                                          | **Notes**                                                                                                                                                               |
|------------------------|------------------------------------|----------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `headerName`           | Column header text                 | Yes                                                |                                                                                                                                                                         |
| `columnIndex`          | Position of the column             | Only when using `ColumnIndexMode.USER_DEFINED` |                                                                                                                                                                         |
| `columnColumnDataType` | Excel cell type for the column     | No                                                 | **If specified, this overrides any `cellTypeMode` setting - even if `cellTypeMode` is set to `AUTO`, the explicitly defined type will be used for this column** |
| `format`               | Format pattern for cell values     | No                                                 | **If specified, this format takes precedence over any automatic formatting from `dataFormatMode`, even when `dataFormatMode` is set to `AUTO_BY_CELL_TYPE`**    |
| `headerStyle`          | Style for this column's header     | No                                                 | overrides class-level style                                                                                                                                             |
| `bodyStyle`            | Style for this column's data cells | No                                                 | overrides class-level style                                                                                                                                             |

**Supported field types:**
- String, Character, char
- Numeric types (Byte, Short, Integer, Long, Float, Double and primitives)
- Boolean/boolean
- Date types (LocalDate, LocalDateTime, Date, java.sql.Date)
- Enum values

> **Note:** Arrays and collection types are not supported.

### @ExcelColumnStyle

The `@ExcelColumnStyle` annotation configures cell styling:

```java
@ExcelColumnStyle(
    cellStyleClass = MyEnumStyle.class,
    enumName = "HEADER_STYLE"
)
```

**Key parameters:**

| Parameter        | Description                                                                                             |
|------------------|---------------------------------------------------------------------------------------------------------|
| `cellStyleClass` | Class or enum that defines the style (must implement `ExcelCellStyle` or extend `CustomExcelCellStyle`) |
| `enumName`       | Enum constant name (when `cellStyleClass` is an enum)                                                   |

## Custom Cell Style

You can define cell styles in two ways:

### 1. Using Enum classes

```java
public enum MyStyles implements ExcelCellStyle {
    HEADER_STYLE(
        PaletteExcelColor.of(ColorPalette.GREY_25_PERCENT),
        DefaultExcelAlign.CENTER_CENTER,
        DefaultExcelBorder.all(ExcelBorderStyle.THICK)
    ),
    BODY_STYLE(
        RgbExcelColor.rgb(255, 255, 255),
        DefaultExcelAlign.GENERAL_CENTER,
        DefaultExcelBorder.all(ExcelBorderStyle.THIN)
    );

    private final ExcelColor backgroundColor;
    private final ExcelAlign align;
    private final ExcelBorder border;

    MyStyles(ExcelColor backgroundColor, ExcelAlign align, ExcelBorder border) {
        this.backgroundColor = backgroundColor;
        this.align = align;
        this.border = border;
    }

    @Override
    public void apply(CellStyle cellStyle) {
        backgroundColor.applyBackground(cellStyle);
        align.applyAlign(cellStyle);
        border.applyAllBorder(cellStyle);
    }
}
```

### 2. Using Custom Classes

```java
public class MyCustomStyle extends CustomExcelCellStyle {
    @Override
    public void configure(ExcelCellStyleConfigurer configurer) {
        configurer.excelColor(PaletteExcelColor.of(ColorPalette.GREY_25_PERCENT));
        configurer.excelAlign(DefaultExcelAlign.CENTER_CENTER);
        configurer.excelBorder(DefaultExcelBorder.all(ExcelBorderStyle.THIN));
    }
}
```

## Modes

The library provides several modes to customize how Excel files are generated.

### CellTypeMode

Controls how cell types are determined:

- `NONE` (default): No automatic cell type determination
- `AUTO`: Automatically determines cell types based on field types

### ColumnIndexMode

Controls how column order is determined:

- `FIELD_ORDER`: Uses the order of field declarations in the class
- `USER_DEFINED`: Uses explicitly defined column indices from `@ExcelColumn.columnIndex`

### DataFormatMode

Controls how data formatting is applied:

- `NONE` (default): No automatic formatting
- `AUTO_BY_CELL_TYPE`: Automatically applies formats based on cell types

### SheetMode

Controls how sheets are created when exporting data:

- `ONE_SHEET`: Creates a single sheet (fails if data exceeds maximum rows)
- `MULTI_SHEET` (default): Creates multiple sheets when data exceeds maximum rows

## Troubleshooting & FAQ

### Common Issues

**Q: Is this library thread-safe?**
A: No. The exporters and underlying Apache POI `Workbook`/`CellStyle` objects are not thread-safe.
Create a new exporter per thread/request and do not share exporter instances across threads.

**Q: Should I use try-with-resources with the exporter?**
A: It depends on your use case:
- **Simple case**: If you don't need `addRows()`, you can chain `build().write()` directly. The `write()` method automatically closes the exporter internally.
```java
SXSSFExporter.builder(MyData.class, data).build().write(outputStream);
```
- **Advanced case**: If you need to call `addRows()` after building, use try-with-resources with the exporter:
```java
try (ExcelExporter<MyData> exporter = SXSSFExporter.builder(MyData.class, data).build()) {
    exporter.addRows(moreData);
    exporter.write(outputStream);
}  // Resources are automatically cleaned up
```

**Q: Numbers are stored as text in Excel instead of numeric values**  
A: You can fix this in two ways:
1. Use `@ExcelColumn(columnColumnDataType = ColumnDataType.NUMBER)` to explicitly set the column type
2. Use `CellTypeMode.AUTO` in the class-level `@Excel` annotation to automatically detect numeric types

**Q: My dates are not formatting correctly in the Excel file**  
A: Make sure you've set the appropriate `format` pattern in your `@ExcelColumn` annotation(ex.`format = "yyyy-MM-dd HH:mm:ss"`) or use `DataFormatMode.AUTO_BY_CELL_TYPE` to apply default date formats.

**Q: How can I format numbers with specific patterns?**  
A: Use the `format` attribute in the `@ExcelColumn` annotation with standard Excel format patterns like `#,##0.00` for numbers or `yyyy-MM-dd` for dates.

**Q: Do I need to specify column indices for all fields?**
A: Only if you're using `ColumnIndexMode.USER_DEFINED`. If you use `FIELD_ORDER` mode, columns will be ordered according to field declaration order in the class.

**Q: Can I reuse an exporter after calling `write()` or `close()`?**
A: No. Once `write()` or `close()` is called, the exporter is closed and cannot be reused. Calling `write()`, `addRows()`, or `close()` after the exporter is closed will throw an `IllegalStateException` (except `close()` which is idempotent). Create a new exporter instance if you need to export again.

## API Documentation

The complete API documentation is available at [Javadoc](https://hee9841.github.io/excel-module/javadoc/).</br>
For the latest version, please check the [GitHub repository](https://github.com/hee9841/excel-module).

## License

This project is licensed under the Apache License 2.0.

## Contact

Project Maintainer - [@hee9841](https://github.com/hee9841)

Project Link: [https://github.com/hee9841/excel-module](https://github.com/hee9841/excel-module)

Issues and Pull Requests: [https://github.com/hee9841/excel-module/issues](https://github.com/hee9841/excel-module/issues) 
