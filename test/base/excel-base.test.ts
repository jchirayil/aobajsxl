// test/base/excel-base.test.ts
import { expect } from 'chai';
import { Excel } from '../../src/index';
import fs from 'fs';
import path from 'path';

describe('Excel', () => {
  let excelBase: Excel;

  beforeEach(() => {
    excelBase = new Excel();
  });

  it('should add a sheet and retrieve its data', async () => {
    const sheetName = 'TestSheet';
    const data = [
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
    ];
    excelBase['setData'](sheetName, data); // Accessing protected method

    const retrievedData = excelBase['getData'](sheetName); // Accessing protected method
    expect(retrievedData).to.deep.equal(data);

    const targetFilePath = path.join(__dirname, '../data/test1-copy.xlsx'); // Construct the target file path
    await excelBase.write(targetFilePath);

    expect(fs.existsSync(targetFilePath)).to.be.true; // Check if the file exists
    fs.unlinkSync(targetFilePath); // Clean up the test file
  });

  it('should update sheet data', () => {
    const sheetName = 'UpdateSheet';
    const initialData = [{ name: 'Charlie', age: 35 }];
    const updatedData = [{ name: 'David', age: 40 }];
    excelBase['setData'](sheetName, initialData);
    excelBase['setData'](sheetName, updatedData); // Accessing protected method

    const retrievedData = excelBase['getData'](sheetName);
    expect(retrievedData).to.deep.equal(updatedData);
  });

  it('should read from an actual XLSX file', async () => {
    const filePath = path.join(__dirname, '../data/test-base.xlsx'); // Construct the file path
    expect(fs.existsSync(filePath)).to.be.true; // Check if the file exists

    await excelBase.read(filePath);

    // Add assertions to verify the data read from the XLSX file
    const sheetData = excelBase['getData']('Sheet1'); // Adjust the sheet name

    expect(sheetData).to.be.an('array');
    // Add more specific assertions based on the content of your test.xlsx file
  });

  it('should read from an actual XLSX file and then save as a new file', async () => {
    const filePath = path.join(__dirname, '../data/test-base.xlsx'); // Construct the file path
    expect(fs.existsSync(filePath)).to.be.true; // Check if the file exists

    await excelBase.read(filePath);

    const targetFilePath = path.join(__dirname, '../data/test-base-copy.xlsx'); // Construct the target file path
    await excelBase.write(targetFilePath);

    expect(fs.existsSync(targetFilePath)).to.be.true; // Check if the file exists
    fs.unlinkSync(targetFilePath); // Clean up the test file
  });

  it('should write to a zip file', async () => {
    const filePath = path.join(__dirname, '../data/test-new-write.xlsx'); // Construct the file path
    excelBase['setData']('test', [{ test: 'value' }]);
    await excelBase.write(filePath);
    expect(fs.existsSync(filePath)).to.be.true;
    fs.unlinkSync(filePath);
  });

  // Add more tests for other methods (process, read, etc.)
});
