// test/base/excel-table.test.ts
import { expect } from 'chai';
import { Excel } from '../../src/index';
import fs from 'fs';
import path from 'path';

describe('Excel - Table', () => {
    let excelBase: Excel;

    beforeEach(() => {
        excelBase = new Excel();
    });

    it('should read from an XLSX file with tables', async () => {
        const filePath = path.join(__dirname, '../data/test-base.xlsx'); // Construct the file path
        expect(fs.existsSync(filePath)).to.be.true; // Check if the file exists

        await excelBase.read(filePath);

        // Add assertions to verify the data read from the XLSX file
        const sheetData = excelBase['getData']('Sheet1'); // Adjust the sheet name

        expect(sheetData).to.be.an('array');
        // Add more specific assertions based on the content of your test.xlsx file

        const writeFilePath = path.join(__dirname, '../data/test-table-base-1.xlsx');
        await excelBase.write(writeFilePath);
        expect(fs.existsSync(writeFilePath)).to.be.true; // Check if the file exists
    });
});