/**
 * Shared Excel helpers for the UpdateUsers table: read table data, apply verify result,
 * write loaded data, and mark load errors.
 * Used by Load Data, Verify, and Update flows.
 */

import {
  getUpdateRowProblemDescription,
} from './verifyUpdateUsers';
import type { VerifyUsersResult } from './verifyUsers';

export type UpdateUsersTableData =
  | { noInputTable: true }
  | { rows: unknown[][]; dataBodyRange: Excel.Range };

export function hasUpdateTableData(
  data: UpdateUsersTableData
): data is { rows: unknown[][]; dataBodyRange: Excel.Range } {
  return !('noInputTable' in data && data.noInputTable);
}

/**
 * Ensures the Update sheet is active (if it exists), loads the UpdateUsers table and its data body values.
 * Call within Excel.run(context => ...).
 */
export async function ensureInputAndGetUpdateUsersTable(
  context: Excel.RequestContext
): Promise<UpdateUsersTableData> {
  const workbook = context.workbook;
  const worksheets = workbook.worksheets;
  const updateSheet = worksheets.getItemOrNullObject('Update');
  const activeSheet = worksheets.getActiveWorksheet();
  activeSheet.load('name');
  await context.sync();

  if (!updateSheet.isNullObject && activeSheet.name !== 'Update') {
    updateSheet.activate();
    await context.sync();
  }

  const table = workbook.tables.getItemOrNullObject('UpdateUsers');
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

/**
 * Applies verify result to the sheet: problem row fill color, clear OK rows,
 * delete all notes and add notes for problem rows.
 * Call within the same Excel.run as ensureInputAndGetUpdateUsersTable.
 */
export async function applyUpdateVerifyResultToSheet(
  context: Excel.RequestContext,
  dataBodyRange: Excel.Range,
  rows: unknown[][],
  result: VerifyUsersResult
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
      const text = getUpdateRowProblemDescription(rows, rowIndex);
      const note = notes.add(firstCells[rowIndex], text);
      note.set({ width: 320, height: 160 });
    }
  }

  await context.sync();
}

/**
 * Writes loaded data values into a single row of the UpdateUsers table.
 * `values` must have the same length as the table columns (20).
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

  // Remove notes that overlap with this row's first cell
  // (notes.items is the full sheet; we delete all and re-add later if needed)
  // For load we just clear the note on this specific row by checking address
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
