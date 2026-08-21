import type { Sheet } from '@fortune-sheet/core';
import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import {
  exportFortuneDataToXlsx,
  parseXlsxToFortuneData,
} from '../src/adapters/xlsx/converter';

function readBlob(blob: Blob): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });
}

function createWorkbook(): ArrayBuffer {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([
    ['Item', 'Amount'],
    ['A', 10],
    ['B', 20],
    ['Total', { t: 'n', f: 'SUM(B2:B3)', v: 30 }],
  ]);
  worksheet['!merges'] = [XLSX.utils.decode_range('A5:B5')];
  worksheet['!cols'] = [{ wch: 18 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Summary');
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
}

describe('XLSX conversion', () => {
  it('imports values, formulas, merges, and dimensions into Fortune-sheet data', () => {
    const [sheet] = parseXlsxToFortuneData(createWorkbook());

    expect(sheet?.name).toBe('Summary');
    expect(sheet?.celldata).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ r: 1, c: 1, v: expect.objectContaining({ v: 10 }) }),
        expect.objectContaining({
          r: 3,
          c: 1,
          v: expect.objectContaining({ f: '=SUM(B2:B3)', v: 30 }),
        }),
      ]),
    );
    expect(sheet?.config?.merge?.['4_0']).toEqual({ r: 4, c: 0, rs: 1, cs: 2 });
    expect(sheet?.config?.columnlen?.['0']).toBeGreaterThan(100);
  });

  it('exports Fortune-sheet data to a valid XLSX workbook', async () => {
    const sheets = parseXlsxToFortuneData(createWorkbook());
    const blob = exportFortuneDataToXlsx(sheets);
    const workbook = XLSX.read(await readBlob(blob), { type: 'array', cellFormula: true });
    const worksheet = workbook.Sheets.Summary;

    expect(blob.type).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    expect(worksheet?.B2?.v).toBe(10);
    expect(worksheet?.B4?.f).toBe('SUM(B2:B3)');
    expect(worksheet?.['!merges']).toEqual([XLSX.utils.decode_range('A5:B5')]);
  });

  it('normalizes duplicate and invalid sheet names on export', async () => {
    const sheets: Sheet[] = [
      { name: '', celldata: [] },
      { name: 'Sheet', celldata: [] },
      { name: 'Sheet', celldata: [] },
    ];
    const workbook = XLSX.read(await readBlob(exportFortuneDataToXlsx(sheets)), {
      type: 'array',
    });

    expect(workbook.SheetNames).toEqual(['Sheet', 'Sheet (2)', 'Sheet (3)']);
  });
});
