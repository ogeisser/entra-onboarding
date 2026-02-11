/**
 * Shared Excel helpers for user tables (Create and Update).
 * Provides parametric functions for reading table data, applying verify results,
 * writing loaded data, and marking load errors.
 */

import type { VerifyUsersResult } from './verifyCore';

// ─── Types ───────────────────────────────────────────────────────────

export type TableData =
  | { noInputTable: true }
  | { rows: unknown[][]; dataBodyRange: Excel.Range };

export function hasTableData(
  data: TableData
): data is { rows: unknown[][]; dataBodyRange: Excel.Range } {
  return !('noInputTable' in data && data.noInputTable);
}

// ─── Read table ──────────────────────────────────────────────────────

/**
 * Ensures the given sheet is active (if it exists), loads the named table and its data body values.
 * Call within Excel.run(context => ...).
 */
export async function ensureInputAndGetTable(
  context: Excel.RequestContext,
  sheetName: string,
  tableName: string
): Promise<TableData> {
  const workbook = context.workbook;
  const worksheets = workbook.worksheets;
  const targetSheet = worksheets.getItemOrNullObject(sheetName);
  const activeSheet = worksheets.getActiveWorksheet();
  activeSheet.load('name');
  await context.sync();

  if (!targetSheet.isNullObject && activeSheet.name !== sheetName) {
    targetSheet.activate();
    await context.sync();
  }

  const table = workbook.tables.getItemOrNullObject(tableName);
  await context.sync();

  if (table.isNullObject) {
    return { noInputTable: true };
  }

  const dataBodyRange = table.getDataBodyRange();
  dataBodyRange.load('values');
  await context.sync();

  const rows = (dataBodyRange.values ?? []) as unknown[][];
  return { rows, dataBodyRange };
}

// ─── Apply verify result ─────────────────────────────────────────────

/**
 * Applies verify result to the sheet: problem row fill color, clear OK rows,
 * delete all notes and add notes for problem rows.
 * Call within the same Excel.run; calls context.sync() at the end.
 *
 * @param getRowProblemFn - Function that returns a human-readable problem description for a row.
 */
export async function applyVerifyResultToSheet(
  context: Excel.RequestContext,
  dataBodyRange: Excel.Range,
  rows: unknown[][],
  result: VerifyUsersResult,
  getRowProblemFn: (rows: unknown[][], rowIndex: number) => string
): Promise<void> {
  const problemSet = new Set(result.problemRowIndices);

  if (result.problemRowIndices.length > 0) {
    for (const rowIndex of result.problemRowIndices) {
      const rowRange = dataBodyRange.getRow(rowIndex);
      rowRange.format.fill.color = '#FFF3CD'; // Light amber for problems
    }
    for (let i = 0; i < rows.length; i++) {
      if (!problemSet.has(i)) {
        dataBodyRange.getRow(i).format.fill.clear();
      }
    }
  } else {
    dataBodyRange.format.fill.clear();
  }

  if (rows.length > 0) {
    const worksheet = dataBodyRange.worksheet;
    const firstCells: Excel.Range[] = [];
    for (let i = 0; i < rows.length; i++) {
      firstCells.push(dataBodyRange.getRow(i).getCell(0, 0));
    }
    const notes = worksheet.notes;
    notes.load('items');
    await context.sync();

    for (const note of notes.items) {
      note.delete();
    }

    for (const rowIndex of result.problemRowIndices) {
      const text = getRowProblemFn(rows, rowIndex);
      const note = notes.add(firstCells[rowIndex], text);
      note.set({ width: 320, height: 160 });
    }
  }

  await context.sync();
}

// ─── Write / mark helpers (used by Update Load Data) ─────────────────

/**
 * Writes loaded data values into a single row of a table.
 * Call within an Excel.run context.
 */
export async function writeLoadedDataToTable(
  context: Excel.RequestContext,
  dataBodyRange: Excel.Range,
  rowIndex: number,
  values: string[]
): Promise<void> {
  const rowRange = dataBodyRange.getRow(rowIndex);
  rowRange.values = [values];
  rowRange.format.fill.clear();

  // Clear any existing note on the first cell of this row
  const worksheet = dataBodyRange.worksheet;
  const firstCell = rowRange.getCell(0, 0);
  const notes = worksheet.notes;
  notes.load('items');
  await context.sync();

  firstCell.load('address');
  await context.sync();

  const cellAddress = firstCell.address;
  for (const note of notes.items) {
    note.cellReference.load('address');
  }
  await context.sync();

  for (const note of notes.items) {
    if (note.cellReference.address === cellAddress) {
      note.delete();
    }
  }

  await context.sync();
}

/**
 * Marks a row as having a load error: sets fill color and adds a note with the error message.
 * Call within an Excel.run context.
 */
export async function markLoadErrorOnRow(
  context: Excel.RequestContext,
  dataBodyRange: Excel.Range,
  rowIndex: number,
  errorMessage: string
): Promise<void> {
  const rowRange = dataBodyRange.getRow(rowIndex);
  rowRange.format.fill.color = '#FFF3CD'; // Light amber

  const firstCell = rowRange.getCell(0, 0);
  const worksheet = dataBodyRange.worksheet;
  const notes = worksheet.notes;

  // Remove existing note on this cell first
  notes.load('items');
  await context.sync();

  firstCell.load('address');
  await context.sync();

  const cellAddress = firstCell.address;
  for (const note of notes.items) {
    note.cellReference.load('address');
  }
  await context.sync();

  for (const note of notes.items) {
    if (note.cellReference.address === cellAddress) {
      note.delete();
    }
  }

  const note = notes.add(firstCell, errorMessage);
  note.set({ width: 320, height: 160 });

  await context.sync();
}
