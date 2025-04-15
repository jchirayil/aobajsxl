// src/base/excel-data-handler.ts
import { ExcelCore } from './excel-core';
import JSZip from 'jszip';

export class ExcelDataHandler extends ExcelCore {
  constructor() {
    super();
  }

  async parseData(zip: JSZip): Promise<void> {
    let _rg0: RegExp | null = null;
    let _rg1: RegExp | null = null;
    let _rs0: RegExpMatchArray[] = [];
    let _rs2: RegExpMatchArray[] | null | undefined;
    let match;
    await (async () => {
      Object.keys(zip.files).map(async (file) => {
        if (!this.schema.hasOwnProperty(file)) {
          this.schema[file] = '';
        }
      });
    })()
      .then(async () => {
        if (this.schema.hasOwnProperty('xl/sharedStrings.xml')) {
          await zip
            .file('xl/sharedStrings.xml')
            ?.async('string')
            .then((data) => {
              _rg0 = /<sst[^]*<\/sst>/gim;
              this.schema['xl/sharedStrings.xml'] = data.replace(_rg0, '{placeholder}');
              _rg1 = /<si><t>([^<]*)<\/t><\/si>/gim;
              _rs0 = [...data.matchAll(_rg1)];
              _rs0.forEach((_r, index) => {
                this.addSharedString(_r[1] || '', index);
              });
            });
        }
      })
      .then(async () => {
        await Promise.all(
          Object.keys(this.schema).map(async (file) => {
            const fileContent = await zip.file(file)?.async('string');
            if (fileContent) {
              switch (file) {
                case 'xl/workbook.xml':
                  _rg0 = /<sheets>([^]*)<\/sheets>/gim;
                  this.schema[file] = fileContent.replace(_rg0, '<sheets>{placeholder}</sheets>');
                  _rg0 = /<sheet\s+[^>]*name="([^"]+)"[^>]*>/g;
                  while ((match = _rg0.exec(fileContent)) !== null) {
                    const sheetName = match[1];
                    const sheetId = parseInt(match[0].match(/sheetId="(\d+)"/)?.[1] || '0', 10);
                    this.addSheet(sheetName, null, sheetId);
                  }
                  break;
                case 'xl/_rels/workbook.xml.rels':
                  _rg0 = /<relationship\s([^]*)\/>/gim;
                  this.schema[file] = fileContent.replace(_rg0, '{placeholder}');
                  _rg0 = /<Relationship\s+[^>]*Target="([^"]+)"[^>]*>/g;
                  while ((match = _rg0.exec(fileContent)) !== null) {
                    const target = match[1];
                    const rId = match[0].match(/Id="([^"]+)"/)?.[1] || '';
                    this.updateSheetTarget(rId, target);
                  }
                  break;
                case '[Content_Types].xml':
                  _rg0 = /<Override([^]*)\/>/gim;
                  this.schema[file] = fileContent.replace(_rg0, '{placeholder}');
                  break;
                default:
                  if (file.includes('xl/worksheets/')) {
                    const matches = file.match(/worksheets\/[^/]+\.xml/);
                    if (matches) {
                      const _sn = matches[0];
                      let _d0: any[] = [];
                      _rg0 = /<sheetData>([^]*)<\/sheetData>/gim;
                      this.schema[file] = fileContent.replace(_rg0, '<sheetData>{placeholder}</sheetData>');
                      _rg0 = /<row\s[^>]*>((<c\s[^>]*>(<v>([^<]*)<\/v>)*<\/c>)*)<\/row>/gim;
                      _rs0 = [...fileContent.matchAll(_rg0)];
                      let _row: { [key: string]: any } = {};
                      _d0 = [];
                      for (let _r = 0; _r < _rs0.length; _r++) {
                        _row = {};
                        _rg0 = /<c\s[^>]*>(<v>([^<]*)<\/v>)*<\/c>/gim;
                        const _brs0 = _rs0[_r][0];
                        _rs2 = [..._brs0.matchAll(_rg0)];
                        for (let _c = 0; _c < _rs2.length; _c++) {
                          const cellMatch = _rs2[_c][0].match(/r="([A-Z]+)(\d+)"/);
                          const _pos = cellMatch ? this.lc(cellMatch[2], cellMatch[1]) : [];
                          const typeMatch = _rs2[_c][0].match(/t="([^"])/);
                          const _t = typeMatch ? typeMatch[1] : null;
                          const valueMatch = _rs2[_c][0].match(/<v>([^<]*)<\/v>/);
                          let _v: string = valueMatch ? valueMatch[1] : '';
                          if (_t === 's' && /^\d+$/.test(_v)) {
                            _v = this.shared[parseInt(_v, 10)];
                          }
                          if (_r === 0) {
                            this.cols.push(_v);
                          } else {
                            _row[this.cols[_pos[1]]] = _v;
                          }
                        }
                        if (_r > 0) {
                          _d0.push(_row);
                        }
                      }
                      this.updateData(_sn, _d0);
                    }
                  } else {
                    this.schema[file] = fileContent;
                  }
                  break;
              }
            }
          })
        );
      });
  }

  async buildData(): Promise<JSZip> {
    const zip = new JSZip();
    await (async () => {
      const _keys = Object.keys(this.schema);
      for (const _k of _keys) {
        if (_k !== 'xl/sharedStrings.xml') {
          let _v: string | null = this.schema[_k];
          if (_v && _v.includes('{placeholder}')) {
            _v = this.updateSchema(_k);
          }
          if (_v) {
            zip.file(_k, _v);
          }
        }
      }
      if (_keys.includes('xl/sharedStrings.xml')) {
        let _v: string | null = this.schema['xl/sharedStrings.xml'];
        if (_v && _v.includes('{placeholder}')) {
          _v = this.updateSchema('xl/sharedStrings.xml');
        }
        if (_v) {
          zip.file('xl/sharedStrings.xml', _v);
        }
      }
    })();
    return zip;
  }

  protected addSheet(
    sheetName: string,
    data: any,
    sheetId: number = 0,
    relationId: string = '',
    target: string = ''
  ): void {
    const _sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].name === sheetName);
    let _rId: string = '';
    let _sheetId: number = 0;
    let _target: string = '';
    if (_sheetKey) {
      _rId = _sheetKey;
      _sheetId = this.sheets[_sheetKey].id;
    } else {
      _sheetId = sheetId < 1 ? Object.keys(this.sheets).length + 1 : sheetId;
      _rId = relationId.length > 0 ? relationId : `rId${_sheetId}`;
    }
    if (this.sheets.hasOwnProperty(_rId)) {
      _sheetId = this.sheets[_rId].id;
      sheetName = this.sheets[_rId].name;
      _target = this.sheets[_rId].target;
    } else {
      _target = target.length > 0 ? target : `worksheets/sheet${_sheetId}.xml`;
    }
    this.sheets[_rId] = { name: sheetName, id: _sheetId, target: _target, data: data };
    this.schema[`xl/${_target}`] =
      `<?xml version="1.0" ?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData>{placeholder}</sheetData></worksheet>`;
  }

  protected updateSheetTarget(relationId: string, target: string): void {
    if (this.sheets.hasOwnProperty(relationId)) {
      this.sheets[relationId].target = target;
    }
  }

  protected updateData(target: string, data: any): void {
    // Find the sheet that matches the target value
    const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].target === target);

    // If a matching sheet is found, update its data property
    if (sheetKey) {
      this.sheets[sheetKey].data = data;
    }
  }

  getSheetNames(): string[] {
    return Object.keys(this.sheets).map((sheetName) => this.sheets[sheetName].name);
  }

  getSheetData(sheetName: string): any {
    const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].name === sheetName);
    if (sheetKey) {
      return this.sheets[sheetKey].data;
    }
    return null;
  }

  setSheetData(sheetName: string, data: any): void {
    this.addSheet(sheetName, data);
  }

  protected updateSchema(key: string): string {
    let _rids: string[] = [];
    let _xml = '';
    let _ret = this.schema[key];
    let _id = 0;
    let _lid = 0;

    if (_ret && _ret.includes('{placeholder}')) {
      switch (key) {
        case 'xl/workbook.xml':
          _rids = Object.keys(this.sheets);
          for (const _rid of _rids) {
            _id = this.sheets[_rid].id;
            _xml += `<sheet name="${this.sheets[_rid].name}" sheetId="${this.sheets[_rid].id}" r:id="${_rid}"/>`;
          }
          _ret = _ret.replace('{placeholder}', _xml);
          break;
        case 'xl/_rels/workbook.xml.rels':
          _rids = Object.keys(this.sheets);
          for (const _rid of _rids) {
            _id = this.sheets[_rid].id;
            if (_id > _lid) {
              _lid = _id;
            }
            _xml += `<Relationship Id="${_rid}" Target="${this.sheets[_rid].target}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>`;
          }
          if (this.schema['xl/sharedStrings.xml']) {
            _lid++;
            _xml += `<Relationship Id="rId${_lid}" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>`;
          }
          if (this.schema['xl/theme/theme1.xml']) {
            _lid++;
            _xml += `<Relationship Id="rId${_lid}" Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>`;
          }
          if (this.schema['xl/styles.xml']) {
            _lid++;
            _xml += `<Relationship Id="rId${_lid}" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>`;
          }
          _ret = _ret.replace('{placeholder}', _xml);
          break;
        case '_rels/.rels':
          if (this.schema['docProps/core.xml']) {
            _xml += `<Relationship Id="rId2" Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>`;
          }
          if (this.schema['docProps/app.xml']) {
            _xml += `<Relationship Id="rId3" Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"/>`;
          }
          _ret = _ret.replace('{placeholder}', _xml);
          break;
        case '[Content_Types].xml':
          _xml += `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;
          _rids = Object.keys(this.sheets);
          for (const _rid of _rids) {
            _xml += `<Override PartName="/xl/${this.sheets[_rid].target}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
          }
          if (this.schema['xl/sharedStrings.xml']) {
            _xml += `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`;
          }
          if (this.schema['xl/theme/theme1.xml']) {
            _xml += `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`;
          }
          if (this.schema['xl/styles.xml']) {
            _xml += `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;
          }
          if (this.schema['docProps/core.xml']) {
            _xml += `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`;
          }
          if (this.schema['docProps/app.xml']) {
            _xml += `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`;
          }
          _ret = _ret.replace('{placeholder}', _xml);
          break;
        case 'xl/sharedStrings.xml':
          _rids = Object.keys(this.shared);
          for (const _rid of _rids.map(Number)) {
            _xml += `<si><t>${this.shared[_rid]}</t></si>`;
          }
          _ret = _ret.replace(
            '{placeholder}',
            `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${_xml}</sst>`
          );
          break;
        default:
          if (key.includes('xl/worksheets/')) {
            const match = key.match(/xl\/(worksheets\/[^/]+\.xml)/);
            if (match) {
              const sheetName = match[1];
              const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].target === sheetName);
              if (sheetKey) {
                _ret = _ret.replace('{placeholder}', this.ws(this.sheets[sheetKey].data));
              }
            }
          }
          break;
      }
    }
    return _ret || '';
  }
}
