/**
 * Shared helper to initialize an Excel sheet with title, description, and table.
 * Used by both Create and Update init handlers.
 */

export interface SheetTemplateConfig {
  sheetName: string;
  tableName: string;
  title: string;
  description: string;
  headers: readonly string[];
  /** Last column letter, e.g. 'S' for 19 columns or 'T' for 20. */
  lastColumnLetter: string;
  columnWidths: [string, number][];
  descriptionRowHeight: number;
}

export async function initSheetTemplate(
  context: Excel.RequestContext,
  config: SheetTemplateConfig
): Promise<void> {
  const {
    sheetName,
    tableName,
    title,
    description,
    headers,
    lastColumnLetter,
    columnWidths,
    descriptionRowHeight,
  } = config;

  const workbook = context.workbook;
  const worksheets = workbook.worksheets;

  // Delete existing sheet if it exists
  const existingSheet = worksheets.getItemOrNullObject(sheetName);
  await context.sync();

  if (!existingSheet.isNullObject) {
    existingSheet.delete();
    await context.sync();
  }

  // Create new sheet
  const sheet = worksheets.add(sheetName);
  await context.sync();

  // Set entire sheet to text format
  const entireSheetRange = sheet.getRange();
  (entireSheetRange as unknown as { numberFormat: string }).numberFormat = '@';

  // Activate the sheet
  sheet.activate();
  await context.sync();

  // Row 1: Title
  const titleCell = sheet.getRange('A1');
  titleCell.values = [[title]];
  const titleRow = sheet.getRange(`A1:${lastColumnLetter}1`);
  titleRow.merge();
  titleRow.format.font.bold = true;
  titleRow.format.font.size = 16;
  titleRow.format.font.color = '#FFFFFF';
  titleRow.format.fill.color = '#0078D4';
  titleRow.format.horizontalAlignment = Excel.HorizontalAlignment.left;
  titleRow.format.verticalAlignment = Excel.VerticalAlignment.center;
  titleRow.format.rowHeight = 25;

  // Row 2: Description
  const descCell = sheet.getRange('A2');
  descCell.values = [[description]];
  const descRow = sheet.getRange(`A2:${lastColumnLetter}2`);
  descRow.merge();
  descRow.format.wrapText = true;
  descRow.format.verticalAlignment = Excel.VerticalAlignment.top;
  descRow.format.font.size = 12;
  descRow.format.rowHeight = descriptionRowHeight;

  // Column widths
  for (const [col, width] of columnWidths) {
    sheet.getRange(col).format.columnWidth = width;
  }

  // Row 3: Headers + Table
  const headerRange = sheet.getRange(`A3:${lastColumnLetter}3`);
  headerRange.values = [headers as unknown as string[]];

  const tables = sheet.tables;
  const newTable = tables.add(`A3:${lastColumnLetter}3`, true);
  newTable.name = tableName;

  await context.sync();
}
