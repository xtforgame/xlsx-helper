import { ParsingOptions, WorkBook, Range } from 'xlsx';
export default class WorkBookReader {
    workBook?: WorkBook;
    readFile(filename: string, opts?: ParsingOptions): WorkBook;
    /** Attempts to parse data */
    read(data: any, opts?: ParsingOptions): WorkBook;
    setWorkBook(workBook: WorkBook): WorkBook;
    forEachRowEx(sheetName: string, cb: (row: any, rowIndex: number, range: Range) => boolean | void, options?: any): Error | void;
    forEachRow(sheetName: string, cb: (row: any[], rowIndex: number, range: Range) => boolean | void): Error | void;
    test(sheetIndex?: number): void;
}
