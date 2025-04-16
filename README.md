# AobaJSXL

AobaJSXL is a TypeScript library designed to parse, manipulate, and generate Excel `.xlsx` files. It provides a robust API to transform JSON data into Excel files and vice versa, leveraging the power of [JSZip](https://stuk.github.io/jszip/) for ZIP file handling and OpenXML standards for Excel file structure.

---

## Features

- **JSON to Excel Transformation**: Easily convert JSON data into `.xlsx` files.
- **Excel Parsing**: Extract and process data from existing `.xlsx` files.
- **Sheet Management**: Add, update, and retrieve data from worksheets.
- **Shared Strings Support**: Efficiently handle shared strings for text reuse in Excel files.
- **Lightweight and Fast**: Built on top of JSZip for efficient ZIP file handling.

---

## Installation

Install AobaJSXL via npm or yarn:

```bash
npm install aobajsxl
```
or
```bash
yarn add aobajsxl
```

---

## Usage

### Importing the Library

The Excel class is the main entry point for interfacing with the library. It provides methods to parse and generate Excel files.

```
import { Excel } from 'aobajsxl';
```

### Parsing an Excel file

```
import { Excel } from 'aobajsxl';

const excel = new Excel();

// Load an Excel file
await excel.read('example.xlsx');

// Access sheet data
const sheetData = excel.getData('Sheet1');
console.log(sheetData);
```

### Generating an Excel file

```
import { Excel } from 'aobajsxl';

const excel = new Excel();

// Add data to a sheet
excel.setData('Sheet1', [
    { Name: 'Alice', Age: 25 },
    { Name: 'Bob', Age: 30 },
]);

// Generate the Excel file
await excel.write('example.xlsx');
```

## API Reference

API Document - [Excel](excel.md)

## Project Structure

```
src/
├── base/
│   ├── excel-core.ts          # Core functionality for handling Excel components
│   ├── excel-data-handler.ts  # Internal class for parsing and generating Excel files
├── index.ts                   # Entry point exposing the Excel class
```

## Development

### Prerequisties

* Node.js (v14 or higher)
* npm or yarn

### Setup

1. Clone the repository:

```bash
git clone https://github.com/jchirayil/aobajsxl.git
cd aobajsxl
```

2. Install dependencies:

```bash
npm install
```

### Build

To build the library

```bash
npm run build
```

### Run Tests

To run the test suite:

```bash
npm test
```

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bugfix.
3. Commit your changes and push the branch.
4. Open a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgement

* [JSZip](https://stuk.github.io/jszip) for Zip file handling.
* OpenXML standards for Excel file structure.

## Contact

For questions or support, please open an issue on the [GitHub repository](https://github.com/jchirayil/aobajsxl/issues).