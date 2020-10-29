import path from 'path';
import XLSX from 'xlsx';

export const getColumnSymbols : (ref : string) => string[] = (ref : string) => {
  const columnCount = XLSX.utils.decode_range(ref).e.c + 1;
  const columnSymbols : string[] = Array.from({ length: columnCount });

  for (let index = 0; index < columnCount; index++) {
    columnSymbols[index] = XLSX.utils.encode_col(index);
  }
  return columnSymbols;
};
