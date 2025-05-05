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
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    fs.writeFileSync(fileName, buffer);
  }

  async readJSON(fileName: string): Promise<any> {
    if (fileName.endsWith('.gz')) {
      const zip = new JSZip();
      const compressedData = fs.readFileSync(fileName);
      const extractedData = await zip.loadAsync(compressedData);
      const fileNames = Object.keys(extractedData.files);
      const jsonFile = extractedData.files[fileNames[0]];
      const jsonContent = await jsonFile.async('string');
      return JSON.parse(jsonContent);
    } else if (fileName.endsWith('.json')) {
      const data = fs.readFileSync(fileName, 'utf-8');
      return JSON.parse(data);
    } else {
      throw new Error('Unsupported file type. Only .json and .gz are supported.');
    }
  }
}
