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

    const targetFilePath = path.join(__dirname, '../data/test-base-copy-1.xlsx'); // Construct the target file path
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

  it('should get sheet names', async () => {
    const filePath = path.join(__dirname, '../data/test-sheets.xlsx'); // Construct the file path
    expect(fs.existsSync(filePath)).to.be.true; // Check if the file exists
    await excelBase.read(filePath);
    const sheetNames = excelBase.getSheetNames();
    expect(sheetNames).to.be.an('array');
    expect(sheetNames).to.include('DepartmentSheet'); // Adjust the expected sheet name
    expect(sheetNames).to.include('CompanySheet'); // Adjust the expected sheet name
    expect(sheetNames).to.include('TransactionSheet'); // Adjust the expected sheet name
  });
  
  it('should read from an actual XLSX file and then save as a new file', async () => {
    const filePath = path.join(__dirname, '../data/test-base.xlsx'); // Construct the file path
    expect(fs.existsSync(filePath)).to.be.true; // Check if the file exists

    await excelBase.read(filePath);

    const targetFilePath = path.join(__dirname, '../data/test-base-copy-2.xlsx'); // Construct the target file path
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

  it('should write formula - simple - to a zip file', async () => {
    const filePath = path.join(__dirname, '../data/test-new-formula-simple.xlsx'); // Construct the file path
    excelBase['setData']('FormulaSheet', [
        {id: 1, operation: 'multiplication', total: '=20*3'},
        {id: 2, operation: 'addition', total: '=5+2+4'}, 
        {id: 3, operation: 'subtraction', total: '=152-24'},
        {id: 4, operation: 'division', total: '=100/5'},
        {id: 5, operation: 'exponentiation', total: '=2^3'},
        {id: 6, operation: 'modulus', total: '=MOD(10,3)'},
        {id: 7, operation: 'square root', total: '=SQRT(16)'},
        {id: 8, operation: 'average', total: '=AVERAGE(1,2,3,4,5)'},
        {id: 9, operation: 'sum', total: '=SUM(1,2,3,4,5)'},
        {id: 10, operation: 'count', total: '=COUNT(1,2,3,4,5)'},
        {id: 11, operation: 'max', total: '=MAX(1,2,3,4,5)'},
        {id: 12, operation: 'min', total: '=MIN(1,2,3,4,5)'},
        {id: 13, operation: 'if', total: '=IF(1>2, "True", "False")'},
        {id: 14, operation: 'concatenate', total: '=CONCATENATE("Hello", " ", "World")'}
    ]);
    await excelBase.write(filePath);
    expect(fs.existsSync(filePath)).to.be.true;
    //fs.unlinkSync(filePath);
  });

  it('should write formula - reference - to a zip file', async () => {
    const filePath = path.join(__dirname, '../data/test-new-formula-ref.xlsx'); // Construct the file path
    excelBase['setData']('FormulaSheet', [
        {id: 1, product: 'Apple', quantity: 12, price: 3, total: '=[@quantity]*[@price]'},
        {id: 2, product: 'Banana', quantity: 12, price: 0.2, total: '=[@quantity]*[@price]'},
        {id: 3, product: 'Cherry', quantity: 5, price: 4, total: '=[@quantity]*[@price]'}   
    ]);
    await excelBase.write(filePath);
    expect(fs.existsSync(filePath)).to.be.true;
    //fs.unlinkSync(filePath);
  });
});
