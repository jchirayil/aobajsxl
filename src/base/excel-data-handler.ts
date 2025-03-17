// src/base/excel-schema-handler.ts
import JSZip from 'jszip';

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
    [key: string]: {
        name: string;
        id: number;
        data: any;
    };
}

export class ExcelDataHandler {
    protected schema: Schema = {
        'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><workbookPr/><sheets>{placeholder}</sheets><definedNames/><calcPr/></workbook>`,
        'xl/_rels/workbook.xml.rels': `<?xml version="1.0" ?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{placeholder}</Relationships>`,
        '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>{placeholder}</Relationships>`,
        '[Content_Types].xml': `<?xml version="1.0" ?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default ContentType="application/xml" Extension="xml"/><Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>{placeholder}</Types>`,
        'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{placeholder}`
    };
    protected shared: SharedStrings = {};
    protected sharedRev: SharedStringsRev = {};
    protected sheets: Sheets = {};
    protected cols: string[] = [];

    constructor() { }

    async parseData(zip: JSZip): Promise<void> {
        let _rg0: RegExp | null = null;
        let _rg1: RegExp | null = null;
        let _rs0: RegExpMatchArray[] = [];
        let _rs1: RegExpMatchArray | null | undefined;
        let _rs2: RegExpMatchArray[] | null | undefined;
        await (async () => {
            Object.keys(zip.files).map(async (file) => {
                if (!this.schema.hasOwnProperty(file)) {
                    this.schema[file] = '';
                }
            });
        })().then(async () => {
            if (this.schema.hasOwnProperty('xl/sharedStrings.xml')) {
                await zip.file('xl/sharedStrings.xml')?.async('string').then((data) => {
                    _rg0 = /<sst[^]*<\/sst>/gim;
                    this.schema['xl/sharedStrings.xml'] = data.replace(_rg0, '{placeholder}');
                    _rg1 = /<si><t>([^<]*)<\/t><\/si>/gim;
                    _rs0 = [...data.matchAll(_rg1)];
                    _rs0.forEach((_r, index) => {
                        this.addSharedString(_r[1] || '', index);
                    });
                });
            }
        }).then(async () => {
            await Promise.all(Object.keys(this.schema).map(async (file) => {
                const fileContent = await zip.file(file)?.async('string');
                if (fileContent) {
                    switch (file) {
                        case 'xl/workbook.xml':
                            _rg0 = /<sheets>([^]*)<\/sheets>/gim;
                            this.schema[file] = fileContent.replace(_rg0, '<sheets>{placeholder}</sheets>');
                            _rs0 = [...fileContent.matchAll(_rg0)];
                            _rg0 = /\bname="([^"]*)"/gim;
                            for (const _r of _rs0) {
                                _rs1 = _r[1]?.match(_rg0);
                                if (_rs1) {
                                    this.addSheet(_rs1[0], null);
                                }
                            }
                            break;
                        case 'xl/_rels/workbook.xml.rels':
                            _rg0 = /<relationship\s([^]*)\/>/gim;
                            this.schema[file] = fileContent.replace(_rg0, '{placeholder}');
                            break;
                        case '[Content_Types].xml':
                            _rg0 = /<Override([^]*)\/>/gim;
                            this.schema[file] = fileContent.replace(_rg0, '{placeholder}');
                            break;
                        default:
                            if (file.includes('xl/worksheets/')) {
                                const matches = file.match(/(?:.*\/)?([^/]+?)(?=\.[^/.]*$)/gim);
                                if (matches) {
                                    const _sn = matches[0];
                                    let _d0: any[] = [];
                                    _rg0 = /<sheetData>([^]*)<\/sheetData>/gim;
                                    this.schema[file] = fileContent.replace(_rg0, '{placeholder}');
                                    _rg0 = /<row\s[^>]*>((<c\s[^>]*>(<v>([^<]*)<\/v>)*<\/c>)*)<\/row>/gim;
                                    _rs0 = [...fileContent.matchAll(_rg0)];
                                    let _row: { [key: string]: any } = {};
                                    for (let _r = 0; _r < _rs0.length; _r++) {
                                        _row = {};
                                        _rg0 = /<c\s[^>]*>(<v>([^<]*)<\/v>)*<\/c>/gim;
                                        const _brs0 = _rs0[_r][0];
                                        _rs2 = [..._brs0.matchAll(_rg0)];
                                        for (let _c = 0; _c < _rs2.length; _c++) {
                                            const cellMatch = _rs2[_c][0].match(/r="(\$?[a-z]{1,3})(\$?[0-9]{1,7})"/gim);
                                            const _pos = cellMatch ? this.lc(cellMatch[0], cellMatch[1]) : [];
                                            const typeMatch = _rs2[_c][0].match(/t="([^"])/gim);
                                            const _t = typeMatch ? typeMatch[0] : null;
                                            const valueMatch = _rs2[_c][0].match(/<v>([^<]*)<\/v>/gim);
                                            let _v: string = valueMatch ? valueMatch[0] : '';
                                            if (_t === 's') { _v = this.shared[_v] };
                                            if (_r === 0) { this.cols.push(_v); } else { _row[this.cols[_pos[1]]] = _v; }
                                        }
                                        if (_r > 0) _d0.push(_row);
                                    }
                                    this.updateData(_sn, _d0);
                                }
                            } else {
                                this.schema[file] = fileContent;
                            }
                            break;
                    }
                }
            }));
        });
    }

    protected addSheet(sheetName: string, data: any, sheetId: number = 0): void {
        const _sheetName = sheetName.toLocaleLowerCase();
        let _sheetId = (sheetId < 1 ? (Object.keys(this.sheets).length + 1) : sheetId);
        if (this.sheets.hasOwnProperty(_sheetName)) { _sheetId = this.sheets[_sheetName].id; }
        this.sheets[_sheetName] = { name: _sheetName, id: _sheetId, data };
        this.schema[`xl/worksheets/${_sheetName}.xml`] = `<?xml version="1.0" ?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData>{placeholder}</sheetData></worksheet>`;
    }

    protected updateData(sheetName: string, data: any, sheetId: number = 0):void{
        this.addSheet(sheetName, data, sheetId);
    }

    protected getSheetData(sheetName: string): any {
        let _sheetName: string | undefined = sheetName.toLocaleLowerCase();
        if (!(_sheetName.length > 0 && this.sheets.hasOwnProperty(_sheetName))) {
            _sheetName = Object.keys(this.sheets)[0];
        }
        return this.sheets[_sheetName].data;
    }

    protected updateSchema(key: string): string {
        let _so: string[] = [];
        let _xml = '';
        let _ret = this.schema[key];
        let _id = 0;
        let _lid = 0;

        if (_ret && _ret.includes('{placeholder}')) {
            switch (key) {
                case 'xl/workbook.xml':
                    _so = Object.keys(this.sheets);
                    for (const sheetName of _so) {
                        _id = this.sheets[sheetName].id;
                        _xml += `<sheet name="${this.sheets[sheetName].name}" sheetId="${_id}" r:id="rId${_id}"/>`;
                    }
                    _ret = _ret.replace('{placeholder}', _xml);
                    break;
                case 'xl/_rels/workbook.xml.rels':
                    _so = Object.keys(this.sheets);
                    for (const sheetName of _so) {
                        _id = this.sheets[sheetName].id;
                        if (_id > _lid) { _lid = _id; }
                        _xml += `<Relationship Id="rId${_id}" Target="worksheets/${sheetName.toLocaleLowerCase()}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>`;
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
                    _so = Object.keys(this.sheets);
                    for (const sheetName of _so) {
                        _xml += `<Override PartName="/xl/worksheets/${sheetName.toLocaleLowerCase()}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
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
                    _so = Object.keys(this.shared);
                    for (const _s of _so) {
                        _xml += `<si><t>${this.shared[_s]}</t></si>`;
                    }
                    _ret = _ret.replace('{placeholder}', `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${_xml}</sst>`);
                    break;
                default:
                    if (key.includes('xl/worksheets/')) {
                        const match = key.match(/(?:.*\/)?([^\/]+?)(?=(?:\.[^\/.]*)?$)/);
                        if (match) {
                            const sheetName = match[1];
                            _ret = _ret.replace('{placeholder}', this.__ws(this.sheets[sheetName].data));
                        }
                    } else {
                        console.log('Error: schema has {placeholder} tag.');
                    }
                    break;
            }
        }
        return _ret || '';
    }

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
}