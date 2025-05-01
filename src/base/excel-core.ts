// src/base/excel-core.ts

interface Schema {
  [key: string]: string | null;
}

interface SharedStrings {
  [key: number]: string;
}

interface SharedStringsRev {
  [key: string]: number;
}

interface Sheets {
  [rId: string]: {
    name: string;
    id: number;
    target: string;
    data: any;
  };
}

export class ExcelCore {
  protected schema: Schema = {
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><workbookPr/><sheets>{placeholder}</sheets><definedNames/><calcPr/></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" ?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{placeholder}</Relationships>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>{placeholder}</Relationships>`,
    '[Content_Types].xml': `<?xml version="1.0" ?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default ContentType="application/xml" Extension="xml"/><Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>{placeholder}</Types>`,
    'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{placeholder}`,
  };
  protected shared: SharedStrings = {};
  protected sharedRev: SharedStringsRev = {};
  protected sheets: Sheets = {};
  protected cols: string[] = [];

  constructor() {}

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

  protected findSharedString(text: string): number {
    const _txt = text.length > 32767 ? text.substring(0, 32766) : text;
    if (this.sharedRev.hasOwnProperty(_txt)) {
      return this.sharedRev[_txt];
    } else {
      return this.addSharedString(_txt);
    }
  }

  protected lc(row: string, col: string): [number, number] {
    const _b = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const _r: [number, number] = [Number.parseInt(row) - 1, 0];
    col = col.toUpperCase();
    for (let i = 0, j = col.length - 1; i < col.length; i++, j--) {
      _r[1] += Math.pow(_b.length, j) * (_b.indexOf(col[i]) + 1);
    }
    _r[1]--;
    return _r;
  }

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

  private header(): string {
    let _headerCells = '';
    const _rowIndex = 1;
    for (let _i = 0; _i < this.cols.length; _i++) {
      _headerCells += `<c r="${this.base(_i, _rowIndex)}" t="s"><v>${this.findSharedString(this.fix(this.cols[_i]))}</v></c>`;
    }
    return `<row r="${_rowIndex}">${_headerCells}</row>`;
  }

  private base(colIndex: number, rowIndex: number): string {
    return `${this.cl(colIndex)}${rowIndex}`;
  }

  private ci(col: string): number {
    let _index = this.cols.indexOf(col);
    if (_index < 0) {
      this.cols.push(col);
      _index = this.cols.length - 1;
    }
    return _index;
  }

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

  private cell(index: number, value: any, rowIndex: number): string {
    let _cell = '';
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
        if (typeof value === 'string' && value.startsWith('=')) {
          _cell = `<c r="${this.base(index, rowIndex)}" ><f>${value.substring(1)}</f></c>`;
        } else {
          _cell = `<c r="${this.base(index, rowIndex)}"><v>${value}</v></c>`;
        }
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

  private fix(text: string): string {
    let _text = text.replace(/^\+/gm, `'+`);
    _text = _text.replace(/[ ]{2,}/gi, ' ');
    _text = _text.replace(/\x08/gi, '');
    //_text = escape(_text.replace(/\t/gi, ' '));
    return _text;
  }

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
