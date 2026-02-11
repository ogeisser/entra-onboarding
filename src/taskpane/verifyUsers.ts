/**
 * Validation for the "CreateUsers" table.
 * Expects rows as 2D array (data rows only, no header).
 * Column indices match USERS_TABLE_HEADERS in App.tsx.
 */

export const UPN_DOMAIN = 'majorel.com';
export const MAIL_DOMAINS = ['majorel.com', 'mj.teleperformance.com'] as const;
export const MAX_DATA_ROWS = 100;

const COL = {
  userPrincipalName: 0,
  mail: 1,
  bmsId: 2,
  localHrId: 3,
  firstName: 5,
  lastName: 6,
  displayName: 7,
  country: 8,
  city: 9,
} as const;

const REQUIRED_COLUMNS: { index: number; name: string }[] = [
  { index: COL.userPrincipalName, name: 'User Principal Name' },
  { index: COL.mail, name: 'Mail' },
  { index: COL.firstName, name: 'First Name' },
  { index: COL.lastName, name: 'Last Name' },
  { index: COL.displayName, name: 'Display Name' },
  { index: COL.country, name: 'Country' },
  { index: COL.city, name: 'City' },
];

/** Spalten, deren Werte eindeutig sein müssen (keine Doppelten). */
const UNIQUE_COLUMNS: { index: number; name: string }[] = [
  { index: COL.userPrincipalName, name: 'User Principal Name' },
  { index: COL.mail, name: 'Mail' },
  { index: COL.bmsId, name: 'BMS ID' },
  { index: COL.localHrId, name: 'Local HR ID' },
];

function cellValue(row: unknown[], index: number): string {
  const raw = row[index];
  return String(raw ?? '').trim();
}

function isValidEmailFormat(value: string): boolean {
  if (!value) return false;
  const parts = value.split('@');
  return parts.length === 2 && parts[0].length > 0 && parts[1].length > 0;
}

function parseEmail(value: string): { local: string; domain: string } | null {
  const parts = value.split('@');
  if (parts.length !== 2 || !parts[0].trim() || !parts[1].trim()) return null;
  return { local: parts[0].trim(), domain: parts[1].trim().toLowerCase() };
}

/** BMS ID: nur Ziffern, keine führende Null (z. B. "0" oder "123" ok, "01" / "007" ungültig). */
function isValidBmsId(value: string): boolean {
  if (value === '') return true;
  return /^(0|[1-9]\d*)$/.test(value);
}

export interface VerifyUsersResult {
  success: boolean;
  totalRows: number;
  okCount: number;
  problemCount: number;
  /** 0-based row indices (in the rows array) that have validation errors */
  problemRowIndices: number[];
  /** true when the CreateUsers table was not found (no create data) */
  noInputTable?: boolean;
}

/** Result to return when the CreateUsers table was not found (no create data). */
export const NO_INPUT_TABLE_VERIFY_RESULT: VerifyUsersResult = {
  success: false,
  totalRows: 0,
  okCount: 0,
  problemCount: 0,
  problemRowIndices: [],
  noInputTable: true,
};

/**
 * Returns human-readable validation error messages for a single row.
 * Mirrors the checks in rowHasValidationErrors.
 */
function getRowValidationErrorMessages(row: unknown[]): string[] {
  const messages: string[] = [];

  for (const { index, name } of REQUIRED_COLUMNS) {
    if (cellValue(row, index) === '') {
      messages.push(`Required field '${name}' is empty`);
    }
  }

  const bmsId = cellValue(row, COL.bmsId);
  const localHrId = cellValue(row, COL.localHrId);
  if (bmsId === '' && localHrId === '') {
    messages.push('BMS ID and Local HR ID are both empty');
  }
  if (bmsId !== '' && !isValidBmsId(bmsId)) {
    messages.push('BMS ID must be a number (digits only, no leading zero)');
  }

  const upn = cellValue(row, COL.userPrincipalName);
  const mail = cellValue(row, COL.mail);

  if (upn) {
    if (!isValidEmailFormat(upn)) {
      messages.push("User Principal Name: invalid format");
    } else {
      const upnParsed = parseEmail(upn);
      if (upnParsed && upnParsed.domain !== UPN_DOMAIN.toLowerCase()) {
        messages.push(`User Principal Name: domain must be ${UPN_DOMAIN}`);
      }
    }
  }

  if (mail) {
    if (!isValidEmailFormat(mail)) {
      messages.push("Mail: invalid format");
    } else {
      const mailParsed = parseEmail(mail);
      const allowed = MAIL_DOMAINS.map((d) => d.toLowerCase());
      if (mailParsed && !allowed.includes(mailParsed.domain)) {
        messages.push(`Mail: domain must be one of ${MAIL_DOMAINS.join(', ')}`);
      }
    }
  }

  if (upn && mail) {
    const upnParsed = parseEmail(upn);
    const mailParsed = parseEmail(mail);
    if (
      upnParsed &&
      mailParsed &&
      upnParsed.local.toLowerCase() !== mailParsed.local.toLowerCase()
    ) {
      messages.push("User Principal Name and Mail local part do not match");
    }
  }

  return messages;
}

/**
 * Returns duplicate-error messages for a row (which unique columns have duplicates).
 */
function getRowDuplicateMessages(rows: unknown[][], rowIndex: number): string[] {
  const messages: string[] = [];
  const row = rows[rowIndex] ?? [];

  for (const { index, name } of UNIQUE_COLUMNS) {
    const value = cellValue(row, index);
    if (value === '') continue;
    const key = value.toLowerCase();
    const indices: number[] = [];
    for (let i = 0; i < rows.length; i++) {
      if (cellValue(rows[i], index).toLowerCase() === key) {
        indices.push(i);
      }
    }
    if (indices.length > 1) {
      messages.push(`Duplicate value in '${name}'`);
    }
  }

  return messages;
}

function rowHasValidationErrors(row: unknown[]): boolean {
  return getRowValidationErrorMessages(row).length > 0;
}

/**
 * Findet Zeilen-Indizes, in denen Werte in den UNIQUE_COLUMNS doppelt vorkommen.
 * Leere Werte werden ignoriert (mehrere leere Zellen gelten nicht als Duplikat).
 */
function findRowsWithDuplicateUniqueValues(rows: unknown[][]): Set<number> {
  const duplicateRowIndices = new Set<number>();

  for (const { index, name: _name } of UNIQUE_COLUMNS) {
    const valueToRowIndices = new Map<string, number[]>();
    for (let i = 0; i < rows.length; i++) {
      const value = cellValue(rows[i], index);
      if (value === '') continue;
      const key = value.toLowerCase();
      const list = valueToRowIndices.get(key) ?? [];
      list.push(i);
      valueToRowIndices.set(key, list);
    }
    for (const indices of valueToRowIndices.values()) {
      if (indices.length > 1) {
        indices.forEach((i) => duplicateRowIndices.add(i));
      }
    }
  }

  return duplicateRowIndices;
}

/**
 * Validates the CreateUsers table data rows.
 * @param rows - 2D array of cell values (data rows only, no header)
 * @returns { success, totalRows, okCount, problemCount, problemRowIndices }
 */
export function verifyUsers(rows: unknown[][]): VerifyUsersResult {
  const totalRows = rows.length;
  const problemRowIndicesSet = new Set<number>();

  for (let i = 0; i < rows.length; i++) {
    if (rowHasValidationErrors(rows[i])) {
      problemRowIndicesSet.add(i);
    }
  }

  const duplicateRows = findRowsWithDuplicateUniqueValues(rows);
  duplicateRows.forEach((i) => problemRowIndicesSet.add(i));

  const problemRowIndices = [...problemRowIndicesSet].sort((a, b) => a - b);
  const problemCount = problemRowIndices.length;
  const okCount = totalRows - problemCount;

  const overMaxRows = totalRows > MAX_DATA_ROWS;
  const success = problemCount === 0 && !overMaxRows;

  return {
    success,
    totalRows,
    okCount,
    problemCount,
    problemRowIndices,
  };
}

const PROBLEM_MESSAGE_SEPARATOR = '\n';

/**
 * Returns a single human-readable problem description for a row, or empty string if the row has no problems.
 * Used e.g. as the content of an Excel note on the row’s first cell.
 */
export function getRowProblemDescription(
  rows: unknown[][],
  rowIndex: number
): string {
  const row = rows[rowIndex];
  if (!row) return '';
  const validation = getRowValidationErrorMessages(row);
  const duplicates = getRowDuplicateMessages(rows, rowIndex);
  const all = [...validation, ...duplicates];
  return all.length === 0 ? '' : all.join(PROBLEM_MESSAGE_SEPARATOR);
}
