/* eslint-disable no-unused-vars, no-undef */

import chai from 'chai';
import { getColumnSymbols, WorkBookReader } from 'library';
import path from 'path';
import XLSX from 'xlsx';

const { expect } = chai;

declare const describe;
declare const it;

describe('Main Test Cases', () => {
  it('WorkBookReader Test 1', () => {
    // const p = path.join(__dirname, '../test-data/big5.xlsx');
    const p = path.join(__dirname, '../test-data/SampleData.xlsx');
    const wb = XLSX.readFile(p);
    const wbr = new WorkBookReader();
    wbr.setWorkBook(wb);
    wbr.test(1);
  });
});
