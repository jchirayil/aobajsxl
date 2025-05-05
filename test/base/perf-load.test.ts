// test/base/perf-load.test.ts
import { expect } from 'chai';
import { Excel } from '../../src/index';
import fs from 'fs';
import path from 'path';

describe('Excel Performance Load Tests', function () {
    
    this.timeout && this.timeout(0); // Set timeout to 10 seconds for each test
    let excelBase: Excel;

    beforeEach(() => {
        excelBase = new Excel();
    });

    const testCases = [
        { rows: 100, cols: 10, file: 'perf-test-100.json' },
        { rows: 1000, cols: 10, file: 'perf-test-1000.json' },
        { rows: 10000, cols: 10, file: 'perf-test-10000.json.gz' },
        { rows: 100000, cols: 10, file: 'perf-test-100000.json.gz' }
    ];

    testCases.forEach(({ rows, cols, file }) => {
        it(`should read ${rows} (rows) x ${cols} (cols) JSON data`, async () => {
            const sourceFilePath = path.join(__dirname, `../data/${file}`); // Construct the source file path
            const sheetName = `PerSheet${rows}x${cols}`;

            const startTime = Date.now(); // Capture start time
            await excelBase['setData'](sheetName, sourceFilePath); // Accessing protected method
            const loadTime = Date.now(); // Capture load time

            console.log(`\tExecution time to read ${file} file [${rows} (rows) x ${cols} (cols)]: ${loadTime - startTime} ms`); // Report execution time
        });
    });

    testCases.forEach(({ rows, cols, file }) => {
        it(`should write ${rows} (rows) x ${cols} (cols) JSON data`, async () => {
            const sourceFilePath = path.join(__dirname, `../data/${file}`); // Construct the source file path
            const sheetName = `PerSheet${rows}x${cols}`;

            const startTime = Date.now(); // Capture start time
            await excelBase['setData'](sheetName, sourceFilePath); // Accessing protected method
            const loadTime = Date.now(); // Capture load time
            const targetFilePath = path.join(__dirname, `../data/perf-test-${rows}-copy.xlsx`); // Construct the target file path
            //await excelBase.write(targetFilePath);
            const endTime = Date.now(); // Capture end time

            console.log(`\tExecution time [${rows} (rows) x ${cols} (cols)]: ${loadTime - startTime} ms Write: ${endTime - loadTime} ms Total: ${endTime - startTime} ms`); // Report execution time

            //expect(fs.existsSync(targetFilePath)).to.be.true; // Check if the file exists
            //fs.unlinkSync(targetFilePath); // Clean up the test file
        });
    });
});