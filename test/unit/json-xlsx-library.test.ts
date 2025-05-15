// tests/unit/json-xlsx-library.spec.ts

import { expect } from 'chai';
// Assuming your library exposes classes or functions like this:
import { JsonParser, JsonGenerator, XlsxParser, XlsxGenerator } from '../../src/index'; // Adjust the import path

describe('JSON Processing', () => {
  let jsonParser: JsonParser;
  let jsonGenerator: JsonGenerator;

  beforeEach(() => {
    jsonParser = new JsonParser();
    jsonGenerator = new JsonGenerator();
  });

  it('should correctly parse a basic JSON object', async () => {
    const jsonString = '{"name": "John Doe", "age": 30, "isEmployed": true}';
    const result = await jsonParser.parse(jsonString);
    expect(result).to.deep.equal({ name: 'John Doe', age: 30, isEmployed: true });
  });

  it('should correctly parse a nested JSON object', async () => {
    const jsonString = '{"address": {"street": "123 Main St", "city": "Anytown"}}';
    const result = await jsonParser.parse(jsonString);
    expect(result).to.deep.equal({ address: { street: '123 Main St', city: 'Anytown' } });
  });

  it('should handle null values in JSON', async () => {
    const jsonString = '{"nullable": null}';
    const result = await jsonParser.parse(jsonString);
    expect(result).to.deep.equal({ nullable: null });
  });

  it('should throw an error for malformed JSON', async () => {
    const jsonString = '{"name": "John Doe", "age": 30,';
    await expect(jsonParser.parse(jsonString)).to.be.rejectedWith(SyntaxError);
  });

  it('should correctly generate a basic JSON object', async () => {
    const data = { name: 'Jane Doe', age: 25 };
    const result = await jsonGenerator.generate(data);
    expect(JSON.parse(result)).to.deep.equal(data);
  });

  it('should correctly generate a nested JSON object', async () => {
    const data = { company: { name: 'Acme Corp', founded: 1900 } };
    const result = await jsonGenerator.generate(data);
    expect(JSON.parse(result)).to.deep.equal(data);
  });
});

describe('XLSX Processing', () => {
  let xlsxParser: XlsxParser;
  let xlsxGenerator: XlsxGenerator;

  // Mocking the JSZip object and its methods for testing without actual file I/O
  const mockZip = (fileData: Record<string, string>) => ({
    file: (name: string) => ({
      async: (type: 'string' | 'arraybuffer' | 'blob' | 'uint8array' | 'nodebuffer') => {
        if (fileData[name]) {
          return Promise.resolve(fileData[name]);
        }
        return Promise.reject(new Error(`File not found: ${name}`));
      },
    }),
    loadAsync: (data: any) => Promise.resolve(mockZip(fileData)),
  } as any);

  beforeEach(() => {
    xlsxParser = new XlsxParser();
    xlsxGenerator = new XlsxGenerator();
  });

  it('should correctly parse a simple XLSX with one sheet', async () => {
    const xlsxData = mockZip({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://.../worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row><c><v>Name</v></c><c><v>Age</v></c></row><row><c><v>John</v></c><c><v>30</v></c></row></sheetData></worksheet>',
      '[Content_Types].xml': '<Types><Override PartName="/xl/workbook.xml" ContentType=".../spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType=".../spreadsheetml.worksheet+xml"/></Types>',
    });
    const result = await xlsxParser.parse(xlsxData);
    expect(result).to.deep.equal({ Sheet1: [{ Name: 'John', Age: '30' }] });
  });

  it('should correctly parse an XLSX with multiple sheets', async () => {
    const xlsxData = mockZip({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://.../worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://.../worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row><c><v>Name</v></c></row><row><c><v>Alice</v></c></row></sheetData></worksheet>',
      'xl/worksheets/sheet2.xml': '<worksheet><sheetData><row><c><v>City</v></c></row><row><c><v>New York</v></c></row></sheetData></worksheet>',
      '[Content_Types].xml': '<Types><Override PartName="/xl/workbook.xml" ContentType=".../spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType=".../spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType=".../spreadsheetml.worksheet+xml"/></Types>',
    });
    const result = await xlsxParser.parse(xlsxData);
    expect(result).to.deep.equal({ Sheet1: [{ Name: 'Alice' }], Sheet2: [{ City: 'New York' }] });
  });

  it('should handle empty cells in XLSX parsing', async () => {
    const xlsxData = mockZip({
      'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://.../worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row><c><v>Name</v></c><c></c></row><row><c></c><c><v>30</v></c></row></sheetData></worksheet>',
      '[Content_Types].xml': '<Types><Override PartName="/xl/workbook.xml" ContentType=".../spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType=".../spreadsheetml.worksheet+xml"/></Types>',
    });
    const result = await xlsxParser.parse(xlsxData);
    expect(result).to.deep.equal({ Sheet1: [{ Name: null, '': null }, { Name: null, '': '30' }] }); // Or however you choose to represent empty cells
  });

  it('should correctly generate a simple XLSX with one sheet', async () => {
    const data = { Sheet1: [{ Name: 'Bob', Age: 40 }] };
    const zip = await xlsxGenerator.generate(data);
    const files = Object.keys(zip.files);
    expect(files).to.include('xl/workbook.xml');
    expect(files).to.include('xl/worksheets/sheet1.xml');
    const sheet1Content = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
    expect(sheet1Content).to.contain('<v>Bob</v>');
    expect(sheet1Content).to.contain('<v>40</v>');
  });

  it('should correctly generate an XLSX with multiple sheets', async () => {
    const data = { Sheet1: [{ Item: 'Apple' }], Sheet2: [{ Price: 1.0 }] };
    const zip = await xlsxGenerator.generate(data);
    const files = Object.keys(zip.files);
    expect(files).to.include('xl/workbook.xml');
    expect(files).to.include('xl/worksheets/sheet1.xml');
    expect(files).to.include('xl/worksheets/sheet2.xml');
    const sheet1Content = await zip.file('xl/worksheets/sheet1.xml')?.async('string');
    expect(sheet1Content).to.contain('<v>Apple</v>');
    const sheet2Content = await zip.file('xl/worksheets/sheet2.xml')?.async('string');
    expect(sheet2Content).to.contain('<v>1.0</v>');
  });
});

// You would need to create actual implementations for JsonParser, JsonGenerator, XlsxParser, and XlsxGenerator
// in your src/index.ts or relevant files.