/**
 * Excel helpers for the CreateUsers table.
 * Thin wrapper around excelTableHelpers with Create-specific sheet/table names.
 */

import { getRowProblemDescription } from './verifyUsers';
import type { VerifyUsersResult } from './verifyCore';
import {
  ensureInputAndGetTable,
  applyVerifyResultToSheet as applyVerifyResultToSheetCore,
  type TableData,
} from './excelTableHelpers';

// Re-export shared types for backward compatibility.
export type { TableData as UsersTableData } from './excelTableHelpers';
export { hasTableData } from './excelTableHelpers';

/**
 * Ensures the Create sheet is active, loads the CreateUsers table and its data body values.
 * Call within Excel.run(context => ...).
 */
export function ensureInputAndGetUsersTable(
  context: Excel.RequestContext
): Promise<TableData> {
  return ensureInputAndGetTable(context, 'Create', 'CreateUsers');
}

/**
 * Applies verify result to the Create sheet.
 * Call within the same Excel.run; calls context.sync() at the end.
 */
export function applyVerifyResultToSheet(
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
    getRowProblemDescription
  );
}
