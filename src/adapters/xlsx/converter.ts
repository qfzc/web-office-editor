import type {
  Cell,
  CellMatrix,
  CellWithRowAndCol,
  Sheet,
  SheetConfig,
} from '@fortune-sheet/core';
import * as XLSX from 'xlsx';

const DEFAULT_ROWS = 100;
const DEFAULT_COLUMNS = 26;

type StyledCell = XLSX.CellObject & {
  s?: unknown;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function colorFromStyle(value: unknown): string | undefined {
  if (!isRecord(value)) return undefined;
  const rgb = value.rgb;
  return typeof rgb === 'string' ? `#${rgb.slice(-6)}` : undefined;
}

function mapHorizontal(value: unknown): number | undefined {
  if (value === 'center') return 0;
  if (value === 'left') return 1;
  if (value === 'right') return 2;
  return undefined;
}

function mapVertical(value: unknown): number | undefined {
  if (value === 'center') return 0;
  if (value === 'top') return 1;
  if (value === 'bottom') return 2;
  return undefined;
}

function toFortuneCell(source: StyledCell): Cell {
  const cell: Cell = {
    v: source.v as Cell['v'],
    m: source.w ?? (source.v === undefined ? '' : String(source.v)),
  };

  if (source.f) cell.f = source.f.startsWith('=') ? source.f : `=${source.f}`;
  if (source.z) cell.ct = { fa: String(source.z), t: source.t };

  if (isRecord(source.s)) {
    const font = isRecord(source.s.font) ? source.s.font : undefined;
    const fill = isRecord(source.s.fill) ? source.s.fill : undefined;
    const alignment = isRecord(source.s.alignment) ? source.s.alignment : undefined;

    if (font?.bold === true) cell.bl = 1;
    if (font?.italic === true) cell.it = 1;
    if (font?.underline === true) cell.un = 1;
    if (typeof font?.name === 'string') cell.ff = font.name;
    if (typeof font?.sz === 'number') cell.fs = font.sz;
    cell.fc = colorFromStyle(font?.color);
    cell.bg = colorFromStyle(fill?.fgColor);
    cell.ht = mapHorizontal(alignment?.horizontal);
    cell.vt = mapVertical(alignment?.vertical);
    if (alignment?.wrapText === true) cell.tb = '2';
  }

  return cell;
}

function createConfig(worksheet: XLSX.WorkSheet): SheetConfig {
  const config: SheetConfig = {};
  const merges = worksheet['!merges'];
  if (merges?.length) {
    config.merge = Object.fromEntries(
      merges.map(({ s, e }) => [
        `${s.r}_${s.c}`,
        { r: s.r, c: s.c, rs: e.r - s.r + 1, cs: e.c - s.c + 1 },
      ]),
    );
  }

  worksheet['!cols']?.forEach((column, index) => {
    const width = column.wpx ?? (column.wch ? column.wch * 8 : undefined);
    if (width) (config.columnlen ??= {})[String(index)] = width;
    if (column.hidden) (config.colhidden ??= {})[String(index)] = 0;
  });

  worksheet['!rows']?.forEach((row, index) => {
    if (!row) return;
    const height = row.hpx ?? (row.hpt ? row.hpt * (96 / 72) : undefined);
    if (height) (config.rowlen ??= {})[String(index)] = height;
    if (row.hidden) (config.rowhidden ??= {})[String(index)] = 0;
  });

  return config;
}

export function parseXlsxToFortuneData(input: ArrayBuffer): Sheet[] {
  const workbook = XLSX.read(input, {
    type: 'array',
    cellDates: false,
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
  });

  return workbook.SheetNames.map((name, index) => {
    const worksheet = workbook.Sheets[name];
    if (!worksheet) throw new Error(`Workbook sheet is missing: ${name}`);

    const range = XLSX.utils.decode_range(worksheet['!ref'] ?? 'A1');
    const celldata: CellWithRowAndCol[] = [];

    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        const source = worksheet[XLSX.utils.encode_cell({ r: row, c: column })] as
          | StyledCell
          | undefined;
        if (!source) continue;
        celldata.push({ r: row, c: column, v: toFortuneCell(source) });
      }
    }

    return {
      name,
      id: `sheet-${index + 1}`,
      order: index,
      status: index === 0 ? 1 : 0,
      hide: workbook.Workbook?.Sheets?.[index]?.Hidden ? 1 : 0,
      row: Math.max(DEFAULT_ROWS, range.e.r + 1),
      column: Math.max(DEFAULT_COLUMNS, range.e.c + 1),
      celldata,
      config: createConfig(worksheet),
    } satisfies Sheet;
  });
}

function matrixFromCells(cells: CellWithRowAndCol[] | undefined): CellMatrix {
  if (!cells?.length) return [];
  const matrix: CellMatrix = [];
  for (const cell of cells) {
    matrix[cell.r] ??= [];
    matrix[cell.r]![cell.c] = cell.v;
  }
  return matrix;
}

function inferCellType(value: Cell['v']): XLSX.ExcelDataType {
  if (typeof value === 'number') return 'n';
  if (typeof value === 'boolean') return 'b';
  return 's';
}

function toSheetJsCell(source: Cell): XLSX.CellObject | undefined {
  if (source.v === undefined && !source.f) return undefined;

  const output: XLSX.CellObject = {
    t: inferCellType(source.v),
    v: source.v ?? '',
  };
  if (source.f) output.f = source.f.replace(/^=/, '');
  if (source.ct?.fa) output.z = source.ct.fa;
  return output;
}

function appendDimensions(worksheet: XLSX.WorkSheet, sheet: Sheet): void {
  if (sheet.config?.columnlen || sheet.config?.colhidden) {
    worksheet['!cols'] = Array.from({ length: sheet.column ?? 0 }, (_, index) => ({
      wpx: sheet.config?.columnlen?.[String(index)],
      hidden: sheet.config?.colhidden?.[String(index)] !== undefined,
    }));
  }
  if (sheet.config?.rowlen || sheet.config?.rowhidden) {
    worksheet['!rows'] = Array.from({ length: sheet.row ?? 0 }, (_, index) => ({
      hpx: sheet.config?.rowlen?.[String(index)],
      hidden: sheet.config?.rowhidden?.[String(index)] !== undefined,
    }));
  }

  const merges = Object.values(sheet.config?.merge ?? {});
  if (merges.length) {
    worksheet['!merges'] = merges.map(({ r, c, rs, cs }) => ({
      s: { r, c },
      e: { r: r + rs - 1, c: c + cs - 1 },
    }));
  }
}

function uniqueSheetName(name: string, used: Set<string>): string {
  const base = (name.trim() || 'Sheet').slice(0, 31);
  let candidate = base;
  let suffix = 2;
  while (used.has(candidate)) {
    const marker = ` (${suffix})`;
    candidate = `${base.slice(0, 31 - marker.length)}${marker}`;
    suffix += 1;
  }
  used.add(candidate);
  return candidate;
}

export function exportFortuneDataToXlsx(sheets: Sheet[]): Blob {
  const workbook = XLSX.utils.book_new();
  const usedNames = new Set<string>();

  for (const sheet of sheets) {
    const worksheet: XLSX.WorkSheet = {};
    const matrix = sheet.data?.length ? sheet.data : matrixFromCells(sheet.celldata);
    let maxRow = 0;
    let maxColumn = 0;

    matrix.forEach((row, rowIndex) => {
      row?.forEach((cell, columnIndex) => {
        if (!cell) return;
        const output = toSheetJsCell(cell);
        if (!output) return;
        worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })] = output;
        maxRow = Math.max(maxRow, rowIndex);
        maxColumn = Math.max(maxColumn, columnIndex);
      });
    });

    for (const { r, c, rs, cs } of Object.values(sheet.config?.merge ?? {})) {
      maxRow = Math.max(maxRow, r + rs - 1);
      maxColumn = Math.max(maxColumn, c + cs - 1);
    }

    worksheet['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: maxColumn },
    });
    appendDimensions(worksheet, sheet);
    XLSX.utils.book_append_sheet(workbook, worksheet, uniqueSheetName(sheet.name, usedNames));
  }

  if (workbook.SheetNames.length === 0) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[]]), 'Sheet1');
  }

  const bytes = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
    compression: true,
    cellStyles: true,
  }) as ArrayBuffer;

  return new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
