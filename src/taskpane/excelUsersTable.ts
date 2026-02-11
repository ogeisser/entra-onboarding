/**
 * Shared Excel helpers for the CreateUsers table: read table data and apply verify result (markings + notes).
 * Used by both Verify and Create flows.
 */

import {
  getRowProblemDescription,
  type VerifyUsersResult,
} from './verifyUsers';

export type UsersTableData =
  | { noInputTable: true }
  | { rows: unknown[][]; dataBodyRange: Excel.Range };

export function hasTableData(
  data: UsersTableData
): data is { rows: unknown[][]; dataBodyRange: Excel.Range } {
  return !('noInputTable' in data && data.noInputTable);
}

/**
 * Ensures the Create sheet is active (if it exists), loads the CreateUsers table and its data body values.
 * Call within Excel.run(context => ...).
 */
export async function ensureInputAndGetUsersTable(
  context: Excel.RequestContext
): Promise<UsersTableData> {
  const workbook = context.workbook;
  const worksheets = workbook.worksheets;
  const createSheet = worksheets.getItemOrNullObject('Create');
  const activeSheet = worksheets.getActiveWorksheet();
  activeSheet.load('name');
  await context.sync();

  if (!createSheet.isNullObject && activeSheet.name !== 'Create') {
    createSheet.activate();
    await context.sync();
  }

  const table = workbook.tables.getItemOrNullObject('CreateUsers');
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
 * Applies verify result to the sheet: problem row fill color, clear OK rows, delete all notes and add notes for problem rows.
 * Call within the same Excel.run as ensureInputAndGetUsersTable; calls context.sync() at the end.
 */
export async function applyVerifyResultToSheet(
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
      const text = getRowProblemDescription(rows, rowIndex);
      const note = notes.add(firstCells[rowIndex], text);
      note.set({ width: 320, height: 160 });
    }
  }

  await context.sync();
}
