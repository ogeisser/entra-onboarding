/**
 * Excel helpers for the UpdateUsers table.
 * Thin wrapper around excelTableHelpers with Update-specific sheet/table names.
 */

import { getUpdateRowProblemDescription } from './verifyUpdateUsers';
import type { VerifyUsersResult } from './verifyCore';
import {
  ensureInputAndGetTable,
  applyVerifyResultToSheet as applyVerifyResultToSheetCore,
  type TableData,
} from './excelTableHelpers';

// Re-export shared types and Update-specific helpers for backward compatibility.
export type { TableData as UpdateUsersTableData } from './excelTableHelpers';
export { hasTableData as hasUpdateTableData } from './excelTableHelpers';
export { writeLoadedDataToTable, markLoadErrorOnRow } from './excelTableHelpers';

/**
 * Ensures the Update sheet is active, loads the UpdateUsers table and its data body values.
 * Call within Excel.run(context => ...).
 */
export function ensureInputAndGetUpdateUsersTable(
  context: Excel.RequestContext
): Promise<TableData> {
  return ensureInputAndGetTable(context, 'Update', 'UpdateUsers');
}

/**
 * Applies verify result to the Update sheet.
 * Call within the same Excel.run; calls context.sync() at the end.
 */
export function applyUpdateVerifyResultToSheet(
  context: Excel.RequestContext,
  dataBodyRange: Excel.Range,
  rows: unknown[][],
  result: VerifyUsersResult
): Promise<void> {
  return applyVerifyResultToSheetCore(
    context,
    dataBodyRange,
    rows,
    result,
    getUpdateRowProblemDescription
  );
}
