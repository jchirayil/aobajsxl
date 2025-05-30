export interface BaseOptions {
    name: string;
    id?: number;
    target?: string; // Optional target for the sheet or table
}

export interface SheetOptions extends BaseOptions {}

export interface TableOptions extends BaseOptions {
    uid?: string; // Unique identifier for the table
}

export interface TableColumn extends BaseOptions {
    uid: string;
}

export interface Table extends BaseOptions {
    id: number; // Unique identifier for the table
    sheetrId: string; // Reference to the sheet this table belongs to
    uid?: string;
    ref?: string;
    columns?: TableColumn[]; // Expecting an array of column names
    rowCount?: number; // Number of rows in the table
}

export interface Tables {
    [rId: string]: Table;
}

export interface Sheet extends BaseOptions {
    tablerIds?: string[];
}

export interface Sheets {
    [rId: string]: Sheet;
}

export interface TableData {
    [rId: string]: any[]; // Expecting an array of rows for each table
}

export interface CellAddress {
    row: number;
    col: number;
}
