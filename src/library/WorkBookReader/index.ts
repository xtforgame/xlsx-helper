import XLSX, {
  ParsingOptions,
  WorkBook,
  Range,
} from 'xlsx';

import { getColumnSymbols } from '../helpers';

export default class WorkBookReader {
  workBook?: WorkBook;

  readFile(filename: string, opts?: ParsingOptions): WorkBook {
    this.workBook = XLSX.readFile(filename, opts);
    return this.workBook;
  }

  /** Attempts to parse data */
  read(data: any, opts?: ParsingOptions): WorkBook {
    this.workBook = XLSX.read(data, opts);
    return this.workBook;
  }

  setWorkBook(workBook: WorkBook): WorkBook {
    this.workBook = workBook;
    return this.workBook;
  }

  forEachRowEx(sheetName : string, cb : (row : any, rowIndex : number, range : Range) => boolean | void, options : any = {}) : Error | void {
    let columnNames : any[] = [];

    return this.forEachRow(sheetName, (row : any[], rowIndex : number, range : Range) => {
      if (rowIndex === range.s.r) {
        columnNames = [...row];
        if (options.getModifiedColumnNames) {
          columnNames = options.getModifiedColumnNames(columnNames);
        }
        return;
      }
      const r : any = {};
      row.forEach((c, i) => r[columnNames[i]] = c);
      return cb(r, rowIndex, range);
    });
  }

  forEachRow(sheetName : string, cb : (row : any[], rowIndex : number, range : Range) => boolean | void) : Error | void {
    const ws = this.workBook!.Sheets[sheetName];
    if (!ws) {
      return new Error(`sheet not found: ${sheetName}`);
    }
    const ref = ws['!ref']!;
    const range = XLSX.utils.decode_range(ref);
    const columnSize = range.e.c - range.s.c + 1;
    for (let r = range.s.r; r <= range.e.r; ++r) {
      const row = Array.from({ length: columnSize });
      for (let c = range.s.c; c <= range.e.c; ++c) {
        const cell_address = { c, r };
        /* if an A1-style address is needed, encode the address */
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        const cell = ws[cell_ref];
        row[c - range.s.c] = cell && cell.v;
      }
      const keepGoing = cb(row, r, range);
      if (keepGoing === false) {
        break;
      }
    }
  }

  test(sheetIndex : number = 0) {
    const wb = this.workBook!;
    console.log('wb.SheetNames :', wb.SheetNames);
    const ws = wb.Sheets[wb.SheetNames[sheetIndex]];
    console.log('ws["!ref"] :', ws['!ref']);

    const ref = ws['!ref']!;
    console.log('ref :', ref);
    const columnSymbols = getColumnSymbols(ref);
    console.log('columnSymbols :', columnSymbols);
    const data = XLSX.utils.sheet_to_json(ws);
    // console.log('data :', data);
    columnSymbols.forEach((columnSymbol) => {

    });

    this.forEachRowEx(wb.SheetNames[sheetIndex], (row) => {
      console.log('row :', row);
    }, {
      getModifiedColumnNames: (cols) => {
        cols[cols.length - 2] = 'sizes';
        cols[cols.length - 1] = 'fits';
        return cols;
      }
    });
  }
}
