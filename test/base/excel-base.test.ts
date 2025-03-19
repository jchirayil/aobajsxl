// test/base/excel-base.test.ts
import { expect } from 'chai';
import { ExcelBase } from '../../src/base/excel-base';
import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';

describe('ExcelBase', () => {
  let excelBase: ExcelBase;

  beforeEach(() => {
    excelBase = new ExcelBase();
  });

  it('should add a sheet and retrieve its data', async() => {
    const sheetName = 'TestSheet';
    const data = [{ name: 'Alice', age: 30 }, { name: 'Bob', age: 25 }];
    excelBase['setData'](sheetName, data); // Accessing protected method

    const retrievedData = excelBase['getData'](sheetName); // Accessing protected method
    expect(retrievedData).to.deep.equal(data);

    const targetFilePath = path.join(__dirname, '../data/test1-copy.xlsx'); // Construct the target file path
    await excelBase.write(targetFilePath);

    expect(fs.existsSync(targetFilePath)).to.be.true; // Check if the file exists

  });

  /*
  it('should update sheet data', () => {
    const sheetName = 'UpdateSheet';
    const initialData = [{ name: 'Charlie', age: 35 }];
    const updatedData = [{ name: 'David', age: 40 }];
    excelBase['addSheet'](sheetName, initialData);
    excelBase['updateData'](sheetName, updatedData); // Accessing protected method

    const retrievedData = excelBase['getSheetData'](sheetName);
    expect(retrievedData).to.deep.equal(updatedData);
  });
*/

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

  });
  

  /*
  it('should write to a zip file', async () => {
    const zip = new JSZip();
    excelBase['addSheet']('test', [{test: 'value'}]);
    await excelBase.process(zip, 'write');
    await excelBase.write(zip, 'write_test.xlsx');
    expect(fs.existsSync('write_test.xlsx')).to.be.true;
    fs.unlinkSync('write_test.xlsx');
  });
*/
  // Add more tests for other methods (process, read, etc.)
});