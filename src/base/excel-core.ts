// src/base/excel-core.ts

interface Schema {
  [key: string]: string;
}

interface SharedStrings {
  [key: number]: string;
}

interface SharedStringsRev {
  [key: string]: number;
}

interface Sheet {
  name: string;
  id: number;
  target: string;
  data: any[]; // Expecting an array of rows
}

interface Sheets {
  [rId: string]: Sheet;
}

const MAX_SHARED_STRING_LENGTH = 32767;
const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

export abstract class ExcelCore {
  protected static readonly XML_TAGS: Schema = {
    WORKBOOK: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><workbookPr/><sheets>{placeholder}</sheets><definedNames/><calcPr/></workbook>`,
    WORKBOOK_RELS: `<?xml version="1.0" ?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{placeholder}</Relationships>`,
    WORKBOOK_RELS_SHARED_STRINGS: `<Relationship Id="rId{placeholder}" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>`,
    WORKBOOK_RELS_THEME: `<Relationship Id="rId{placeholder}" Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>`,
    WORKBOOK_RELS_STYLES: `<Relationship Id="rId{placeholder}" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>`,
    RELS: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>{placeholder}</Relationships>`,
    RELS_CORE: `<Relationship Id="rId2" Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>`,
    RELS_APP: `<Relationship Id="rId3" Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"/>`,
    CONTENT_TYPES: `<?xml version="1.0" ?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default ContentType="application/xml" Extension="xml"/><Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>{placeholder}</Types>`,
    CONTENT_TYPE_PART_WORKBOOK: `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
    CONTENT_TYPE_PART_WORKSHEET: `<Override PartName="/xl/{placeholder}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    CONTENT_TYPE_PART_SHARED_STRINGS: `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`,
    CONTENT_TYPE_PART_THEME: `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`,
    CONTENT_TYPE_PART_STYLES: `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
    CONTENT_TYPE_PART_CORE: `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
    CONTENT_TYPE_PART_APP: `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`,
    SHARED_STRINGS: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{placeholder}`,
    SHARED_STRING_LIST: `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{placeholder}</sst>`,
    WORKSHEET: `<?xml version="1.0" ?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData>{placeholder}</sheetData></worksheet>`,
  };

  protected schema: Schema = {
    'xl/workbook.xml': ExcelCore.XML_TAGS.WORKBOOK,
    'xl/_rels/workbook.xml.rels': ExcelCore.XML_TAGS.WORKBOOK_RELS,
    '_rels/.rels': ExcelCore.XML_TAGS.RELS,
    '[Content_Types].xml': ExcelCore.XML_TAGS.CONTENT_TYPES,
    'xl/sharedStrings.xml': ExcelCore.XML_TAGS.SHARED_STRINGS,
  };
  protected shared: SharedStrings = {};
  protected sharedRev: SharedStringsRev = {};
  protected sheets: Sheets = {};
  protected cols: string[] = [];

  constructor() {}

  /**
   * Adds a string to the shared strings table.
   * @param text The string to add.
   * @param index Optional index to insert the string at. If -1, appends to the end.
   * @returns The index of the added/existing shared string.
   */
  protected addSharedString(text: string, index: number = -1): number {
    if (index < 0) {
      index = Object.keys(this.shared).length;
    }
    if (this.shared.hasOwnProperty(index)) {
      index++;
      return this.addSharedString(text, index);
    } else {
      this.shared[index] = text;
      this.sharedRev[text] = index;
      return index;
    }
  }

  /**
   * Finds or adds a string to the shared strings table.
   * Limits the string length to MAX_SHARED_STRING_LENGTH.
   * @param text The string to find or add.
   * @returns The index of the shared string.
   */
  protected findSharedString(text: string): number {
    const _txt = text.length > MAX_SHARED_STRING_LENGTH ? text.substring(0, MAX_SHARED_STRING_LENGTH - 1) : text;
    if (this.sharedRev.hasOwnProperty(_txt)) {
      return this.sharedRev[_txt];
    } else {
      return this.addSharedString(_txt);
    }
  }

  /**
   * Converts an Excel cell reference (e.g., "A1") to its column and row index (0-based).
   * @param row The row part of the reference (e.g., "1").
   * @param col The column part of the reference (e.g., "A").
   * @returns An array containing the row and column index: [rowIndex, colIndex].
   */
  protected lc(row: string, col: string): [number, number] {
    const _r: [number, number] = [Number.parseInt(row) - 1, 0];
    col = col.toUpperCase();
    for (let i = 0, j = col.length - 1; i < col.length; i++, j--) {
      _r[1] += Math.pow(ALPHABET.length, j) * (ALPHABET.indexOf(col[i]) + 1);
    }
    _r[1]--;
    return _r;
  }

  /**
   * Generates the XML for a worksheet's data.
   * Handles large datasets by processing rows incrementally.
   * @param data An array of row objects.
   * @returns The XML string for the <sheetData> element.
   */
  protected ws(data: any[]): string {
    let _data = '';
    this.cols = [];
    try {
      for (let _index = 0; _index < data.length; _index++) {
        _data += this.row(data[_index], _index);
      }
      _data = this.header() + _data;
    } catch (ex) {
      console.log('Exception _ws:', ex, ' data length:', _data.length);
    }
    return _data;
  }

  /**
   * Generates the XML for a single row.
   * @param row An object representing a row of data.
   * @returns The XML string for the <row> element.
   */
  protected row(row: { [key: string]: any }, index: number): string {
    const _rowIndex = index + 2;
    let _rowCells = '';
    let _colIndex = 0;
    for (const _key of Object.keys(row)) {
      _colIndex = this.ci(_key);
      _rowCells += this.cell(_colIndex, row[_key], _rowIndex);
      _colIndex++;
    }
    return `<row r="${_rowIndex}">${_rowCells}</row>`;
  }

  /**
   * Generates the XML for the header row.
   * @returns The XML string for the header <row> element.
   */
  private header(): string {
    let _headerCells = '';
    const _rowIndex = 1;
    for (let _i = 0; _i < this.cols.length; _i++) {
      _headerCells += `<c r="${this.base(_i, _rowIndex)}" t="s"><v>${this.findSharedString(this.fix(this.cols[_i]))}</v></c>`;
    }
    return `<row r="${_rowIndex}">${_headerCells}</row>`;
  }

  /**
   * Generates the Excel cell reference (e.g., "A1").
   * @param colIndex The 0-based column index.
   * @param rowIndex The 1-based row index.
   * @returns The Excel cell reference string.
   */
  private base(colIndex: number, rowIndex: number): string {
    return `${this.cl(colIndex)}${rowIndex}`;
  }

  /**
   * Gets the column index for a given column name.
   * If the column name doesn't exist, it's added to the `cols` array.
   * @param col The column name.
   * @returns The 0-based column index.
   */
  private ci(col: string): number {
    let _index = this.cols.indexOf(col);
    if (_index < 0) {
      this.cols.push(col);
      _index = this.cols.length - 1;
    }
    return _index;
  }

  /**
   * Converts a 0-based column index to its Excel column letter(s).
   * @param index The 0-based column index.
   * @returns The Excel column letter(s).
   */
  private cl(index: number): string {
    if (typeof index !== 'number') {
      return '';
    }
    const _prefix = Math.floor(index / 26);
    const _letter = String.fromCharCode(97 + (index % 26)).toUpperCase();
    if (_prefix === 0) {
      return _letter;
    }
    return this.cl(_prefix - 1) + _letter;
  }

  /**
   * Generates the XML for a single cell.
   * @param index The 0-based column index.
   * @param value The cell value.
   * @param rowIndex The 1-based row index.
   * @returns The XML string for the <c> element.
   */
  private cell(index: number, value: any, rowIndex: number): string {
    let _cell = '';
    let _cellValue = '';
    let _type = this.type(value);
    if (value === undefined || value === null) {
      _type = 'null';
      value = '';
    }
    switch (_type) {
      case 'bool':
        value = value.toLocaleLowerCase() === 'true';
      case 'boolean':
        _cell = `<c r="${this.base(index, rowIndex)}" t="b"><v>${value}</v></c>`;
        break;
      case 'date':
        _cell = `<c r="${this.base(index, rowIndex)}" t="d"><v>${value}</v></c>`;
        break;
      case 'formula':
        value = value.replace(/\[@([^\]]+)\]/g, (_: any, match: string) => {
          return this.base(this.ci(match), rowIndex);
        });
        _cellValue =
          typeof value === 'string' && value.startsWith('=') ? `<f>${value.substring(1)}</f>` : `<v>${value}</v>`;
        _cell = `<c r="${this.base(index, rowIndex)}">${_cellValue}</c>`;
        break;
      case 'null':
      case 'string':
        _cell = `<c r="${this.base(index, rowIndex)}" t="s"><v>${this.findSharedString(this.fix(value))}</v></c>`;
        break;
      case 'number':
        _cell = `<c r="${this.base(index, rowIndex)}" t="n"><v>${value}</v></c>`;
        break;
      default:
        _cell = `<c r="${this.base(index, rowIndex)}"><v>${value}</v></c>`;
        break;
    }
    return _cell;
  }

  /**
   * Determines the data type of a value for Excel.
   * @param val The value to check.
   * @returns The determined data type string.
   */
  private type(val: any): string {
    let _type = 'string';
    if ([true, false].includes(val)) {
      _type = 'boolean';
    } else if (['TRUE', 'FALSE', 'True', 'False', 'true', 'false'].includes(val)) {
      _type = 'bool';
    } else if (val != null && !isNaN(val)) {
      _type = 'number';
    } else if (typeof val === 'string' && val.startsWith('=')) {
      _type = 'formula';
    }
    return _type;
  }

  /**
   * Fixes certain characters in text to prevent issues in Excel.
   * @param text The text to fix.
   * @returns The fixed text.
   */
  private fix(text: string): string {
    let _text = text.replace(/^\+/gm, `'+`);
    _text = _text.replace(/[ ]{2,}/gi, ' ');
    _text = _text.replace(/\x08/gi, '');
    //_text = escape(_text.replace(/\t/gi, ' '));
    return _text;
  }

  /**
   * Flattens a nested JavaScript object into a single-level object.
   * Keys of nested properties are concatenated with '-'.
   * This method might not be suitable for very large or deeply nested objects due to its recursive nature and string concatenation.
   * Consider alternative approaches for large, complex data.
   * @param obj The object to flatten.
   * @returns A flattened object.
   */
  private flatten(obj: any): any {
    const _obj1 = JSON.parse(JSON.stringify(obj));
    const _obj2 = JSON.parse(JSON.stringify(obj));

    const __propCheck = (obj: any, key: string): boolean => {
      return obj[key]?.hasOwnProperty(key) && typeof obj[key] === 'object' && obj[key] != null;
    };

    if (typeof obj === 'object') {
      for (const _k1 in _obj2) {
        if (__propCheck(_obj2, _k1)) {
          delete _obj1[_k1];
          for (const _k2 in _obj2[_k1]) {
            if (_obj2[_k1].hasOwnProperty(_k2)) {
              _obj1[_k1 + '-' + _k2] = _obj2[_k1][_k2];
            }
          }
        }
      }
      const _hasObject = Object.keys(_obj1).some((_k) => __propCheck(_obj1, _k));
      if (_hasObject) {
        return this.flatten(_obj1);
      }
    }
    return _obj1;
  }
}
