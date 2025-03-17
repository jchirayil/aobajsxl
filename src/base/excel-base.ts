// src/base/excel-base.ts
import { ExcelZipHandler } from './excel-zip-handler';
import { ExcelDataHandler } from './excel-data-handler';
import { ExcelSheetHandler } from './excel-sheet-handler';
import JSZip from 'jszip';

export class ExcelBase extends ExcelSheetHandler {
  private zipHandler: ExcelZipHandler;
  private dataHandler: ExcelDataHandler;

  constructor() {
    super();
    this.zipHandler = new ExcelZipHandler();
    this.dataHandler = new ExcelDataHandler();
  }

  async read(fileName: string): Promise<void> {
    const zip = await this.zipHandler.readZip(fileName);
    await this.dataHandler.parseData(zip);
  }

  async write(zip: JSZip, filename: string): Promise<void> {
    await this.zipHandler.writeZip(zip, filename);
  }

  async process(zip: JSZip, actionType: string = 'write'): Promise<void> {
    if (actionType === 'read') {
      await this.dataHandler.parseData(zip);
      // ... use sheetHandler and schemaHandler to process data
    } else {
      // ... use sheetHandler and schemaHandler to update data
      // ... use zipHandler to write the zip file
    }
  }

}