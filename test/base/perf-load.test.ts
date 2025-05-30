// test/base/perf-load.test.ts
import { describe, expect, beforeEach, it } from 'vitest';
import { Excel } from '../../src/index.js';
import fs from 'fs';
import path from 'path';

describe('Excel Performance Load Tests', function () {
    //this.timeout && this.timeout(0); // Set timeout to 10 seconds for each test
    let excelBase: Excel;

    beforeEach(() => {
        excelBase = new Excel();
    });

    const testCases = [
        { rows: 100, cols: 10, file: 'perf-test-100.json', write: true },
        { rows: 1000, cols: 10, file: 'perf-test-1000.json', write: true },
        { rows: 10000, cols: 10, file: 'perf-test-10000.json.gz', write: true },
        { rows: 100000, cols: 10, file: 'perf-test-100000.json.gz', write: false },
    ];

    testCases.forEach(({ rows, cols, file }) => {
        it(`should read ${rows} (rows) x ${cols} (cols) JSON data`, async () => {
            const sourceFilePath = path.join(__dirname, `../data/${file}`); // Construct the source file path
            const sheetName = `PerSheet${rows}x${cols}`;

            const startTime = Date.now(); // Capture start time
            await excelBase['setData'](sheetName, sourceFilePath, 'table1'); // Accessing protected method
            const loadTime = Date.now(); // Capture load time
            console.log(
                `\tExecution time to read [${rows} (rows) x ${cols} (cols)] from ${file}: ${loadTime - startTime} ms`
            ); // Report execution time

            expect(excelBase['getData'](sheetName, 'table1')?.length).toEqual(rows); // Check if the data length matches the expected rows
        });
    });

    testCases.forEach(({ rows, cols, file, write }) => {
        if (write) {
            it(`should write ${rows} (rows) x ${cols} (cols) JSON data`, async () => {
                const sourceFilePath = path.join(__dirname, `../data/${file}`); // Construct the source file path
                const sheetName = `PerSheet${rows}x${cols}`;

                await excelBase['setData'](sheetName, sourceFilePath); // Accessing protected method
                const loadTime = Date.now(); // Capture load time
                const targetFilePath = path.join(__dirname, `../data/perf-test-${rows}-copy.xlsx`); // Construct the target file path
                await excelBase.write(targetFilePath);
                const endTime = Date.now(); // Capture end time
                console.log(
                    `\tExecution time to write [${rows} (rows) x ${cols} (cols)] to ${targetFilePath}: ${endTime - loadTime} ms`
                ); // Report execution time

                expect(fs.existsSync(targetFilePath)).to.be.true; // Check if the file exists
                //fs.unlinkSync(targetFilePath); // Clean up the test file
            });
        }
    });
});
