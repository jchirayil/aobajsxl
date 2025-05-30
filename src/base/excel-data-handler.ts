// src/base/excel-data-handler.ts
import { ExcelCore } from './excel-core.js';
import JSZip from 'jszip';
import {
    SheetOptions,
    Sheet,
    Sheets,
    Table,
    TableData,
    TableOptions,
    Tables,
    CellAddress,
    TableColumn,
} from './excel-shared.js';
import { table } from 'console';

const REGEX = {
    SHARED_STRINGS: /<si><t>([^<]*)<\/t><\/si>/gim,
    SHARED_STRINGS_TABLE: /<sst[^]*<\/sst>/gim,
    SHEETS: /<sheets>([^]*)<\/sheets>/gim,
    SHEET_NAME: /<sheet\s+[^>]*name="([^"]+)"[^>]*>/g,
    SHEET_ID: /sheetId="(\d+)"/,
    RELATIONSHIP: /<relationship\s([^]*)\/>/gim,
    RELATIONSHIP_TARGET: /<Relationship\s+[^>]*Target="([^"]+)"[^>]*>/g,
    RELATIONSHIP_ID: /Id="([^"]+)"/,
    ROW: /<row\s[^>]*>((<c\s[^>]*>(<v>([^<]*)<\/v>)*<\/c>)*)<\/row>/gim,
    CELL: /<c\s[^>]*>(<v>([^<]*)<\/v>)*<\/c>/gim,
    CELL_REF: /r="([A-Z]+)(\d+)"/,
    CELL_TYPE: /t="([^"])/,
    CELL_VALUE: /<v>([^<]*)<\/v>/,
    SHEET_DATA: /<sheetData>([^]*)<\/sheetData>/gim,
    TABLE_PARTS: /<tableParts[^>]*>[\s\S]*?<\/tableParts>/gim,
    PLACEHOLDER: /\{placeholder\}/im,
    CONTENT_PART: /<Override([^]*)\/>/gim,
    WORKSHEET: /worksheets\/[^/]+\.xml/,
    WORKSHEET_PATH: /xl\/(worksheets\/[^/]+\.xml)/,
    WORKSHEET_RELS: /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/,
    WORKSHEET_REL_NAME: /^xl\/worksheets\/_rels\/(sheet\d+)\.xml\.rels$/,
    TABLE_PATH: /^xl\/tables\/table\d+\.xml$/,
    TABLE_TARGET: /^xl\/(tables\/table\d+\.xml)$/,
    TABLE_COLUMNS: /<tableColumn[^>]*id="([^"]+)"[^>]*xr3:uid="([^"]+)"[^>]*name="([^"]+)"\/>/g,
};

export class ExcelDataHandler extends ExcelCore {
    readonly PLACEHOLDER = '{placeholder}';
    protected sheets: Sheets = {};
    protected tables: Tables = {};
    protected data: TableData = {};
    constructor() {
        super();
    }

    /**
     * Asynchronously parses data from the Excel file (JSZip object).
     * Processes shared strings and sheet data incrementally to handle large files.
     * @param zip The JSZip object representing the Excel file.
     */
    async parseData(zip: JSZip): Promise<void> {
        // 1. Build schema: initialize all files in zip.files as keys in this.schema
        Object.keys(zip.files).forEach((file) => {
            if (!this.schema.hasOwnProperty(file)) {
                this.schema[file] = '';
            }
        });

        // 2. Define ordered handlers for each file type
        const handlers = [
            {
                pattern: (f: string) => f === 'xl/sharedStrings.xml',
                handler: (file: string, fileContent: string) => {
                    this.schema[file] = fileContent.replace(REGEX.SHARED_STRINGS_TABLE, this.PLACEHOLDER);
                    const _rs0 = [...fileContent.matchAll(REGEX.SHARED_STRINGS)];
                    _rs0.forEach((_r, index) => {
                        this.addSharedString(_r[1] || '', index);
                    });
                },
            },
            {
                pattern: (f: string) => REGEX.WORKSHEET_RELS.test(f),
                handler: (file: string, fileContent: string) => {
                    this.schema[file] = fileContent.replace(REGEX.RELATIONSHIP, this.PLACEHOLDER);
                },
            },
            {
                pattern: (f: string) => REGEX.TABLE_PATH.test(f),
                handler: (file: string, fileContent: string) => {
                    this.schema[file] = ExcelCore.XML_TAGS.TABLE;
                    // Extract table attributes using regex
                    const tableMatch = fileContent.match(
                        /<table[^>]*id="([^"]+)"[^>]*xr:uid="([^"]+)"[^>]*name="([^"]+)"[^>]*displayName="([^"]+)"[^>]*ref="([^"]+)"[^>]*totalsRowShown="([^"]+)"/
                    );
                    const tableInfo = tableMatch
                        ? {
                              id: Number(tableMatch[1]),
                              uid: tableMatch[2],
                              name: tableMatch[3],
                              ref: tableMatch[5],
                          }
                        : {};
                    if (tableInfo.id != undefined) {
                        const _rid: string = this.addTable(
                            '',
                            { id: tableInfo.id, name: tableInfo.name, uid: tableInfo.uid },
                            []
                        );
                        // Extract tableColumn entries
                        const tableColumns: TableColumn[] = [];
                        let colMatch;
                        while ((colMatch = REGEX.TABLE_COLUMNS.exec(fileContent)) !== null) {
                            tableColumns.push({
                                id: Number(colMatch[1]),
                                uid: colMatch[2],
                                name: colMatch[3],
                            });
                        }
                        this.tables[_rid].columns = tableColumns;
                    }
                },
            },
            {
                pattern: (f: string) => REGEX.WORKSHEET_PATH.test(f),
                handler: (file: string, fileContent: string) => {
                    const _wsMatches = file.match(REGEX.WORKSHEET);
                    if (_wsMatches) {
                        const _sn = _wsMatches[0];
                        let _d0: any[] = [];
                        let _c0: string[] = [];
                        this.schema[file] = fileContent
                            .replace(REGEX.SHEET_DATA, `<sheetData>${this.PLACEHOLDER}</sheetData>`)
                            .replace(REGEX.TABLE_PARTS, `<tableParts>${this.PLACEHOLDER}</tableParts>`);
                        const _rs0 = [...fileContent.matchAll(REGEX.ROW)];
                        let _row: { [key: string]: any } = {};
                        _d0 = [];
                        _c0 = [];
                        for (let _r = 0; _r < _rs0.length; _r++) {
                            _row = {};
                            const _brs0 = _rs0[_r][0];
                            const _rs2 = [..._brs0.matchAll(REGEX.CELL)];
                            for (let _c = 0; _c < _rs2.length; _c++) {
                                const cellMatch = _rs2[_c][0].match(REGEX.CELL_REF);
                                const _pos = cellMatch ? this.lc(cellMatch[2], cellMatch[1]) : [];
                                const typeMatch = _rs2[_c][0].match(REGEX.CELL_TYPE);
                                const _t = typeMatch ? typeMatch[1] : null;
                                const valueMatch = _rs2[_c][0].match(REGEX.CELL_VALUE);
                                let _v: string = valueMatch ? valueMatch[1] : '';
                                if (_t === 's' && /^\d+$/.test(_v)) {
                                    _v = this.shared[parseInt(_v, 10)];
                                }
                                if (_r === 0) {
                                    _c0.push(_v);
                                } else {
                                    _row[_c0[_pos[1]]] = _v;
                                }
                            }
                            if (_r > 0) _d0.push(_row);
                        }
                        this.updateData(_sn, _d0);
                    }
                },
            },
            {
                pattern: (f: string) => f === 'xl/workbook.xml',
                handler: (file: string, fileContent: string) => {
                    this.schema[file] = fileContent.replace(REGEX.SHEETS, `<sheets>${this.PLACEHOLDER}</sheets>`);
                    let match;
                    while ((match = REGEX.SHEET_NAME.exec(fileContent)) !== null) {
                        const sheetName = match[1];
                        const sheetId = parseInt(match[0].match(REGEX.SHEET_ID)?.[1] || '0', 10);
                        this.addSheet({ name: sheetName, id: sheetId }, { name: '' }, []);
                    }
                },
            },
            {
                pattern: (f: string) => f === 'xl/_rels/workbook.xml.rels',
                handler: (file: string, fileContent: string) => {
                    this.schema[file] = fileContent.replace(REGEX.RELATIONSHIP, this.PLACEHOLDER);
                    let match;
                    while ((match = REGEX.RELATIONSHIP_TARGET.exec(fileContent)) !== null) {
                        const target = match[1];
                        const rId = match[0].match(REGEX.RELATIONSHIP_ID)?.[1] || '';
                        this.updateSheetTarget(rId, target);
                    }
                },
            },
            {
                pattern: (f: string) => f === '_rels/.rels',
                handler: (file: string, fileContent: string) => {
                    this.schema[file] = fileContent.replace(REGEX.RELATIONSHIP, this.PLACEHOLDER);
                },
            },
            {
                pattern: (f: string) => f === f, // fallback for any other file
                handler: (file: string, fileContent: string) => {
                    this.schema[file] = fileContent;
                },
            },
        ];

        // 3. Sort keys by handler order
        const keys = Object.keys(this.schema).sort((a, b) => {
            const aIdx = handlers.findIndex((h) => h.pattern(a));
            const bIdx = handlers.findIndex((h) => h.pattern(b));
            return aIdx - bIdx;
        });

        // 4. Process each file in order
        for (const file of keys) {
            const fileContent = await zip.file(file)?.async('string');
            if (!fileContent) continue;
            const handler = handlers.find((h) => h.pattern(file));
            if (handler) {
                handler.handler(file, fileContent);
            }
        }
    }

    async buildData(): Promise<JSZip> {
        const zip = new JSZip();
        await (async () => {
            // Define a sort order function for the keys
            const order = (k: string): number => {
                if (k === 'xl/workbook.xml') return 0;
                if (k === 'xl/_rels/workbook.xml.rels') return 1;
                if (k === '_rels/.rels') return 2;
                if (/^xl\/worksheets\/sheet\d+\.xml$/.test(k)) return 3;
                if (/^xl\/tables\/table\d+\.xml$/.test(k)) return 4;
                if (/^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(k)) return 5;
                if (k === 'xl/sharedStrings.xml') return 99;
                return 10;
            };

            const _keys = Object.keys(this.schema).sort((a, b) => order(a) - order(b));

            for (const _k of _keys) {
                let _v: string | null = this.schema[_k];
                if (_v && _v.includes(this.PLACEHOLDER)) _v = this.updateSchema(_k);
                if (_v) {
                    zip.file(_k, _v);
                }
            }
        })();
        return zip;
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
            console.log('data:', data);
            // this.sheets[sheetKey].data = data;
        }
    }

    protected updateSchema(key: string): string {
        let _ret = this.schema[key];

        //pattern-handler pairs
        const _keyHandlers = [
            {
                pattern: (f: string) => f === 'xl/workbook.xml',
                handler: () => {
                    let _xml: string = '';
                    for (const _rid of Object.keys(this.sheets)) {
                        _xml += `<sheet name="${this.sheets[_rid].name}" sheetId="${this.sheets[_rid].id ?? 0}" r:id="${_rid}"/>`;
                    }
                    _ret = _ret.replace(this.PLACEHOLDER, _xml);
                },
            },
            {
                pattern: (f: string) => f === 'xl/_rels/workbook.xml.rels',
                handler: () => {
                    let _xml: string = '';
                    let _lid: number = 0;
                    for (const _rid of Object.keys(this.sheets)) {
                        const _id: number = this.sheets[_rid].id ?? 0;
                        if (_id > _lid) _lid = _id;
                        _xml += `<Relationship Id="${_rid}" Target="${this.sheets[_rid].target}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>`;
                    }
                    if (this.schema['xl/sharedStrings.xml']) {
                        _lid++;
                        _xml += ExcelCore.XML_TAGS.WORKBOOK_RELS_SHARED_STRINGS.replace(
                            this.PLACEHOLDER,
                            _lid.toString()
                        );
                    }
                    if (this.schema['xl/theme/theme1.xml']) {
                        _lid++;
                        _xml += ExcelCore.XML_TAGS.WORKBOOK_RELS_THEME.replace(this.PLACEHOLDER, _lid.toString());
                    }
                    if (this.schema['xl/styles.xml']) {
                        _lid++;
                        _xml += ExcelCore.XML_TAGS.WORKBOOK_RELS_STYLES.replace(this.PLACEHOLDER, _lid.toString());
                    }
                    _ret = _ret.replace(this.PLACEHOLDER, _xml);
                },
            },
            {
                pattern: (f: string) => f === '_rels/.rels',
                handler: () => {
                    let _xml: string = '';
                    if (this.schema['docProps/core.xml']) _xml += ExcelCore.XML_TAGS.RELS_CORE;
                    if (this.schema['docProps/app.xml']) _xml += ExcelCore.XML_TAGS.RELS_APP;
                    _ret = _ret.replace(this.PLACEHOLDER, _xml);
                },
            },
            {
                pattern: (f: string) => f === '[Content_Types].xml',
                handler: () => {
                    let _xml: string = ExcelCore.XML_TAGS.CONTENT_TYPE_PART_WORKBOOK;
                    //adding worksheets
                    for (const _rid of Object.keys(this.sheets)) {
                        _xml += ExcelCore.XML_TAGS.CONTENT_TYPE_PART_WORKSHEET?.replace(
                            this.PLACEHOLDER,
                            this.sheets[_rid].target || ''
                        );
                    }
                    //adding tables
                    for (const _rid of Object.keys(this.tables)) {
                        _xml += ExcelCore.XML_TAGS.CONTENT_TYPE_PART_TABLE?.replace(
                            this.PLACEHOLDER,
                            this.tables[_rid].target || ''
                        );
                    }
                    if (this.schema['xl/sharedStrings.xml'])
                        _xml += ExcelCore.XML_TAGS.CONTENT_TYPE_PART_SHARED_STRINGS;
                    if (this.schema['xl/theme/theme1.xml']) _xml += ExcelCore.XML_TAGS.CONTENT_TYPE_PART_THEME;
                    if (this.schema['xl/styles.xml']) _xml += ExcelCore.XML_TAGS.CONTENT_TYPE_PART_STYLES;
                    if (this.schema['docProps/core.xml']) _xml += ExcelCore.XML_TAGS.CONTENT_TYPE_PART_CORE;
                    if (this.schema['docProps/app.xml']) _xml += ExcelCore.XML_TAGS.CONTENT_TYPE_PART_APP;
                    _ret = _ret.replace(REGEX.PLACEHOLDER, _xml);
                },
            },
            {
                pattern: (f: string) => f === 'xl/sharedStrings.xml',
                handler: () => {
                    let _xml: string = '';
                    for (const _rid of Object.keys(this.shared).map(Number)) {
                        _xml += `<si><t>${this.shared[_rid]}</t></si>`;
                    }
                    _ret = _ret.replace(
                        this.PLACEHOLDER,
                        ExcelCore.XML_TAGS.SHARED_STRING_LIST?.replace(this.PLACEHOLDER, _xml)
                    );
                },
            },
            {
                pattern: (f: string) => REGEX.WORKSHEET_PATH.test(f),
                handler: () => {
                    const match = key.match(REGEX.WORKSHEET_PATH);
                    if (match) {
                        const sheetName = match[1];
                        let _sheetData: string = '';
                        let _tableParts: string = '';
                        const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].target === sheetName);
                        if (sheetKey && this.sheets[sheetKey].tablerIds) {
                            _tableParts = `<tableParts count="${this.sheets[sheetKey].tablerIds?.length}">`;
                            let _tableCount: number = 0;
                            let _prevTableRef: CellAddress[] = [];
                            this.sheets[sheetKey].tablerIds?.forEach((tableId: string, idx: number) => {
                                if (this.tables[tableId] && this.data[tableId]) {
                                    // If this is not the first table, adjust its .ref property
                                    if (idx > 0) {
                                        const _currentTabelRef: CellAddress[] = this.parseAddress(
                                            this.tables[tableId].ref ?? 'A1:A1'
                                        );
                                        if (!_currentTabelRef[1]) {
                                            _currentTabelRef.push({
                                                row: _currentTabelRef[0].row,
                                                col: _currentTabelRef[0].col,
                                            });
                                        }
                                        if (_currentTabelRef[0].col < _prevTableRef[1].col + 2)
                                            _currentTabelRef[0].col = _prevTableRef[1].col + 2;

                                        this.tables[tableId].ref = this.generateAddress(_currentTabelRef);
                                    }
                                    _sheetData += this.ws(this.tables[tableId], this.data[tableId]);
                                    _prevTableRef = this.parseAddress(this.tables[tableId].ref ?? 'A1');
                                    _tableCount++;
                                    _tableParts += `<tablePart r:id="rId${_tableCount}"/>`;
                                }
                            });
                            _tableParts += '</tableParts>';
                            _ret = _ret.replace(
                                `<sheetData>${this.PLACEHOLDER}</sheetData>`,
                                `<sheetData>${_sheetData}</sheetData>`
                            );
                            _ret = _ret.replace(`<tableParts>${this.PLACEHOLDER}</tableParts>`, _tableParts);
                        }
                    }
                },
            },
            {
                pattern: (f: string) => REGEX.WORKSHEET_RELS.test(f),
                handler: () => {
                    const match = key.match(REGEX.WORKSHEET_REL_NAME);
                    if (match) {
                        const sheetName = `worksheets/${match[1]}.xml`;
                        const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].target === sheetName);
                        if (sheetKey && this.sheets[sheetKey].tablerIds) {
                            let _xml: string = '';
                            let _rid: number = 1;
                            this.sheets[sheetKey].tablerIds?.forEach((tableId) => {
                                if (this.tables[tableId]) {
                                    _xml += `<Relationship Id="rId${_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../${this.tables[tableId].target || ''}"/>`;
                                    _rid++;
                                }
                            });
                            _ret = _ret.replace(this.PLACEHOLDER, _xml);
                        }
                    }
                },
            },
            {
                pattern: (f: string) => REGEX.TABLE_PATH.test(f),
                handler: () => {
                    const match = key.match(REGEX.TABLE_TARGET);
                    if (match) {
                        const tableName = match[1];
                        const tableKey = Object.keys(this.tables).find((key) => this.tables[key].target === tableName);
                        if (tableKey && this.tables[tableKey]) {
                            let _xml: string = '';
                            let _xmlCols: string = '';
                            let _colCount: number = 0;
                            this.tables[tableKey].columns?.forEach((column) => {
                                _xmlCols += `<tableColumn id="${column.id}" xr3:uid="${column.uid}" name="${column.name}"/>`;
                                _colCount++;
                            });
                            _xmlCols = `<tableColumns count="${_colCount}">${_xmlCols}</tableColumns>`;
                            _xml = `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" id="${this.tables[tableKey].id}" xr:uid="${this.tables[tableKey].uid}" name="${this.tables[tableKey].name}" displayName="${this.tables[tableKey].name}" ref="${this.tables[tableKey].ref}" totalsRowShown="0">`;
                            _xml += `<autoFilter ref="${this.tables[tableKey].ref}" xr:uid="${this.tables[tableKey].uid}"/>${_xmlCols}<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
                            _ret = _ret.replace(this.PLACEHOLDER, _xml);
                        }
                    }
                },
            },
        ];

        if (_ret && REGEX.PLACEHOLDER.test(_ret)) {
            const _keyParser = _keyHandlers.find((h) => h.pattern(key));
            if (_keyParser) {
                _keyParser.handler();
            } else {
                console.log('Pattern-Handler not found for key:', key);
            }
        }
        return _ret || '';
    }

    getSheet(sheetName: string): Sheet | null {
        const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].name === sheetName);
        if (sheetKey) {
            return this.sheets[sheetKey];
        }
        return null;
    }

    getTable(sheetName: string, tableName: string): Table | null {
        const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].name === sheetName);
        if (sheetKey && this.sheets[sheetKey].tablerIds) {
            const tableId = this.sheets[sheetKey].tablerIds?.find((id) => this.tables[id].name === tableName);
            if (tableId && this.tables[tableId]) {
                return this.tables[tableId];
            }
        }
        return null;
    }

    getSheetNames(): string[] {
        return Object.keys(this.sheets).map((sheetName) => this.sheets[sheetName].name);
    }

    getTalbeNames(sheetName: string = ''): string[] {
        if (sheetName !== '') {
            const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].name === sheetName);
            if (sheetKey) {
                return (this.sheets[sheetKey].tablerIds ?? []).map((tableId) => this.tables[tableId].name);
            }
        } else {
            return Object.keys(this.tables).map((tableId) => this.tables[tableId].name);
        }
        return [];
    }

    getData(sheetName: string, tableName: string = ''): any[] | null {
        console.log('data-handler getData:', sheetName, tableName);
        const sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].name === sheetName);
        if (sheetKey) {
            let tableKey: string | undefined = '';
            if (tableName) {
                tableKey = Object.keys(this.tables).find((key) => this.tables[key].name === tableName);
            } else {
                tableKey = this.sheets[sheetKey].tablerIds?.[0] || undefined;
            }
            //const _tableName = tableName || this.sheets[sheetKey].tablerIds?.[0] || '';
            //if (_tableName) {
            //    const tableKey = Object.keys(this.tables).find((key) => this.tables[key].name === _tableName);
            console.log(
                'getData:',
                sheetName,
                sheetKey,
                tableName,
                this.sheets[sheetKey].tablerIds,
                tableKey,
                this.tables
            );
            if (tableKey && this.data[tableKey]) {
                return this.data[tableKey];
            }
            //}
        }
        return null;
    }

    setData(sheetName: string, tableName: string = '', data: any[]): void {
        const sheetOption: SheetOptions = { name: sheetName };
        const tableOption: TableOptions = { name: tableName };
        this.addSheet(sheetOption, tableOption, data);
    }

    check(): void {
        //console.log('sheets:', this.sheets, '\ntables:', this.tables, '\nOver');
    }

    private addSheet(sheetOptions: SheetOptions, tableOptions: TableOptions = { name: '' }, data: any[]): void {
        const _sheetKey = Object.keys(this.sheets).find((key) => this.sheets[key].name === sheetOptions.name);
        let _rId: string = '';
        let _sheetId: number = 0;
        let _target: string = '';
        let _name: string = '';
        let _rel: string = '';
        if (_sheetKey) {
            _rId = _sheetKey;
            _sheetId = this.sheets[_sheetKey].id || 0;
        } else {
            _sheetId = sheetOptions.id && sheetOptions.id > 0 ? sheetOptions.id : Object.keys(this.sheets).length + 1;
            _name = sheetOptions.name ?? 'Sheet1';
            _rId = `rId${_sheetId}`;
            _target = `worksheets/sheet${_sheetId}.xml`;
            _rel = `worksheets/_rels/sheet${_sheetId}.xml.rels`;
        }
        if (this.sheets.hasOwnProperty(_rId)) {
            _sheetId = this.sheets[_rId].id || 0;
            _name = this.sheets[_rId].name;
            _target = this.sheets[_rId].target || '';
            _rel = `worksheets/_rels/${_name}.xml.rels`;
        }
        this.sheets[_rId] = { name: _name, id: _sheetId, target: _target };
        this.schema[`xl/${_target}`] = ExcelCore.XML_TAGS.WORKSHEET;
        this.schema[`xl/${_rel}`] = ExcelCore.XML_TAGS.WORKBOOK_RELS;
        this.addTable(_rId, tableOptions, data);
    }

    private addTable(sheetId: string, tableOptions: TableOptions, data: any[]): string {
        const _table: Table = {
            id: tableOptions.id || Object.keys(this.tables).length + 1,
            name: tableOptions.name || `table${tableOptions.id || Object.keys(this.tables).length + 1}`,
            sheetrId: sheetId,
            uid: tableOptions.uid || `{${this.generateGUID()}}`,
            ref: 'A1:A1',
            target: `tables/table${tableOptions.id || Object.keys(this.tables).length + 1}.xml`,
        };

        let _rId: string = `rId${_table.id}`;

        //if (tableOptions.name === '') {
        //    _table.id = Object.keys(this.tables).length + 1;
        //    _table.name = `table${_table.id}`;
        //    _rId = `rId${_table.id}`;
        //    _table.target = `tables/table${_table.id}.xml`;
        //}
        const _tableKey = Object.keys(this.tables).find((key) => this.tables[key].name === _table.name);
        if (_tableKey) {
            _rId = _tableKey;
            _table.id = this.tables[_tableKey].id;
            _table.name = this.tables[_tableKey].name;
            _table.target = this.tables[_tableKey].target || `tables/table${this.tables[_tableKey].id}.xml`;
            _table.uid = this.tables[_tableKey].uid || `{${this.generateGUID()}}`;
            _table.ref = this.tables[_tableKey].ref || '';
        }

        this.tables[_rId] = JSON.parse(JSON.stringify(_table)); // deep copy
        this.schema[`xl/${_table.target}`] = ExcelCore.XML_TAGS.TABLE;
        this.addData(_rId, data);

        // Update the reference in the sheet
        if (sheetId.length > 0 && this.sheets[sheetId]) {
            if (!this.sheets[sheetId].tablerIds) {
                this.sheets[sheetId].tablerIds = [];
            }
            if (!this.sheets[sheetId].tablerIds?.includes(_rId)) {
                this.sheets[sheetId].tablerIds?.push(_rId);
            }
        }
        console.log('add table after:', this.tables, '\nsheet:', this.sheets);
        return _rId;
    }

    private addData(tableId: string, data: any[]): void {
        this.data[tableId] = data;
    }
}
