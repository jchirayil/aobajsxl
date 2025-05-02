# Changelog

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