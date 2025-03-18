// src/base/excel-zip-handler.ts
import JSZip from 'jszip';
import fs from 'fs';

export class ExcelZipHandler {
  async readZip(fileName: string): Promise<JSZip> {
    const zip = new JSZip();
    await zip.loadAsync(fs.readFileSync(fileName));
    return zip;
  }

  async writeZip(zip: JSZip, fileName: string): Promise<void> {
    const buffer = await zip.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 9 },
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    fs.writeFileSync(fileName, buffer);
  }
}