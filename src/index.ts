// src/base/excel-base.ts
import { ExcelZipHandler } from './base/excel-zip-handler';
import { ExcelDataHandler } from './base/excel-data-handler';
import JSZip from 'jszip';

/**
 * Excel class to handle reading and writing Excel files.
 * @class
 * @summary This class provides methods to read and write Excel files.
 * @example
 * const excel = new Excel();
 * await excel.read('example.xlsx');
 * const data = excel.getData('Sheet1');
 * console.log(data);
 * await excel.write('output.xlsx');
 * const sheetNames = excel.getSheetNames();
 * console.log(sheetNames);
 */
export class Excel {
  private zipHandler: ExcelZipHandler;
  private dataHandler: ExcelDataHandler;

  constructor() {
    this.zipHandler = new ExcelZipHandler();
    this.dataHandler = new ExcelDataHandler();
  }

  /**
   * Reads an Excel file and parses its data.
   * @param fileName The name of the file to read.
   * @returns A promise that resolves when the file is read.
   * @summary Reads an Excel file and parses its data.
   * @example
   * const excel = new Excel();
   * await excel.read('example.xlsx');
   * const data = excel.getData('Sheet1');
   * console.log(data);
   */
  async read(fileName: string): Promise<void> {
    const zip = await this.zipHandler.readZip(fileName);
    await this.dataHandler.parseData(zip);
  }

  /**
   * Writes the data to an Excel file.
   * @param filename The name of the file to write.
   * @returns A promise that resolves when the file is written.
   * @summary Writes the data to an Excel file.
   * @example
   * const excel = new Excel();
   * excel.setData('TestSheet', [{ name: 'Alice', age: 30 },{ name: 'Bob', age: 25 }]);
   * await excel.write('output.xlsx');
   */
  async write(filename: string): Promise<void> {
    const zip = await this.dataHandler.buildData();
    await this.zipHandler.writeZip(zip, filename);
  }

  /**
   * Gets the names of all sheets in the Excel file.
   * @returns An array of sheet names.
   * @summary Gets the names of all sheets in the Excel file.
   * @example
   * const excel = new Excel();
   * await excel.read('example.xlsx');
   * const sheetNames = excel.getSheetNames();
   * console.log(sheetNames);
   */
  getSheetNames(): string[] {
    return this.dataHandler.getSheetNames();
  }

  /**
   * Gets the data from a specific sheet.
   * @param sheetName The name of the sheet to get data from.
   * @returns An array of JSON data from the specified sheet.
   * @summary Gets the JSON data from a specific sheet.
   * @example
   * const excel = new Excel();
   * await excel.read('example.xlsx');
   * const data = excel.getData('Sheet1');
   * console.log(data);
   */
  getData(sheetName: string): any[] {
    return this.dataHandler.getSheetData(sheetName);
  }

  /**
   * Sets the data for a specific sheet.
   * @param sheetName The name of the sheet to set data for.
   * @param data An array of JSON data to set in the specified sheet.
   * @summary Sets the JSON data for a specific sheet.
   * @example
   * const excel = new Excel();
   * excel.setData('TestSheet', [{ name: 'Alice', age: 30 },{ name: 'Bob', age: 25 }]);
   */
  setData(sheetName: string, data: any[]): void {
    this.dataHandler.setSheetData(sheetName, data);
  }

  protected async process(zip: JSZip, actionType: string = 'write'): Promise<void> {
    if (actionType === 'read') {
      await this.dataHandler.parseData(zip);
      // ... use sheetHandler and schemaHandler to process data
    } else {
      // ... use sheetHandler and schemaHandler to update data
      // ... use zipHandler to write the zip file
    }
  }
}
