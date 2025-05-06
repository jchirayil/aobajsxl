# Changelog

## [1.0.3] - 2025-05-06

### Summary
This release focuses on performance improvements and optimizations for handling large datasets. Key enhancements include optimizations to shared string handling, batch processing for large data operations.

### Features

- **Support for reading JSON data from files**
  - The `setData` method now supports reading JSON data directly from files.
  - Supported file formats:
    - `.json`: Standard JSON files.
    - `.json.gz`: Compressed JSON files with GZIP compression.
  - Examples: reading JSON data from files:

```typescript
await excelBase.setData('Sheet1', '/path/to/data.json');
await excelBase.setData('Sheet2', '/path/to/data.json.gz');
```

## [1.0.2] - 2025-05-01

### Summary
This release introduces **Supporting Functions** for handling formulas and improves formula parsing capabilities. It also includes examples for better understanding.

### Features
- **Supporting Functions**: Added support for formulas (starting with `=`).
  - Example:
    ```json
    [
      { "id": 1, "operation": "multiplication", "total": "=20*3" },
      { "id": 2, "operation": "addition", "total": "=5+2+4" }
    ]
    ```
    Result:
    ```
    id   operation        total
    1    multiplication   60
    2    addition         11
    ```

- **Formula Parsing**: Automatically resolves cell references in formulas.
  - Formulas must follow specific syntax (e.g., `=[@ColumnA] + [@ColumnB]`).
  - Example:
    ```json
    [
      { "id": 1, "product": "apple", "quantity": 50, "price": 3, "total": "=[@quantity]*[@price]" },
      { "id": 2, "product": "banana", "quantity": 10, "price": 1.4, "total": "=[@quantity]*[@price]" }
    ]
    ```
    Result:
    ```
    id   product   quantity   price   total
    1    apple           50       3     150
    2    banana          10     1.4      14
    ```

### Limitations
- Complex formulas with external references are not supported.

For more details, see the [documentation](https://github.com/jchirayil/aobajsxl).