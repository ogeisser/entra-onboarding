/**
 * Validation for the "UpdateUsers" table.
 * Expects rows as 2D array (data rows only, no header).
 * Column indices match UPDATE_TABLE_HEADERS in App.tsx (Object ID at index 0, then same as Create).
 */

import {
  UPN_DOMAIN,
  MAIL_DOMAINS,
  MAX_DATA_ROWS,
  type VerifyUsersResult,
} from './verifyUsers';

/** Column indices for the UpdateUsers table (Object ID prepended). */
const COL = {
  objectId: 0,
  userPrincipalName: 1,
  mail: 2,
  bmsId: 3,
  localHrId: 4,
  firstName: 6,
  lastName: 7,
  displayName: 8,
  country: 9,
  city: 10,
} as const;

const REQUIRED_COLUMNS: { index: number; name: string }[] = [
  { index: COL.objectId, name: 'Object ID' },
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
  { index: COL.objectId, name: 'Object ID' },
  { index: COL.userPrincipalName, name: 'User Principal Name' },
  { index: COL.mail, name: 'Mail' },
  { index: COL.bmsId, name: 'BMS ID' },
  { index: COL.localHrId, name: 'Local HR ID' },
];

const UUID_REGEX =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function cellValue(row: unknown[], index: number): string {
  const raw = row[index];
  return String(raw ?? '').trim();
}

function isValidEmailFormat(value: string): boolean {
  if (!value) return false;
  const parts = value.split('@');
  return parts.length === 2 && parts[0].length > 0 && parts[1].length > 0;
}

function parseEmail(
  value: string
): { local: string; domain: string } | null {
  const parts = value.split('@');
  if (parts.length !== 2 || !parts[0].trim() || !parts[1].trim()) return null;
  return { local: parts[0].trim(), domain: parts[1].trim().toLowerCase() };
}

/** BMS ID: nur Ziffern, keine führende Null. */
function isValidBmsId(value: string): boolean {
  if (value === '') return true;
  return /^(0|[1-9]\d*)$/.test(value);
}

/** Result to return when the UpdateUsers table was not found (no update data). */
export const NO_INPUT_TABLE_UPDATE_VERIFY_RESULT: VerifyUsersResult = {
  success: false,
  totalRows: 0,
  okCount: 0,
  problemCount: 0,
  problemRowIndices: [],
  noInputTable: true,
};

/**
 * Returns human-readable validation error messages for a single row of the UpdateUsers table.
 */
function getRowValidationErrorMessages(row: unknown[]): string[] {
  const messages: string[] = [];

  // Required fields
  for (const { index, name } of REQUIRED_COLUMNS) {
    if (cellValue(row, index) === '') {
      messages.push(`Required field '${name}' is empty`);
    }
  }

  // Object ID must be a valid UUID
  const objectId = cellValue(row, COL.objectId);
  if (objectId !== '' && !UUID_REGEX.test(objectId)) {
    messages.push('Object ID must be a valid UUID');
  }

  // BMS ID or Local HR ID must be present
  const bmsId = cellValue(row, COL.bmsId);
  const localHrId = cellValue(row, COL.localHrId);
  if (bmsId === '' && localHrId === '') {
    messages.push('BMS ID and Local HR ID are both empty');
  }
  if (bmsId !== '' && !isValidBmsId(bmsId)) {
    messages.push('BMS ID must be a number (digits only, no leading zero)');
  }

  // UPN format & domain
  const upn = cellValue(row, COL.userPrincipalName);
  const mail = cellValue(row, COL.mail);

  if (upn) {
    if (!isValidEmailFormat(upn)) {
      messages.push('User Principal Name: invalid format');
    } else {
      const upnParsed = parseEmail(upn);
      if (upnParsed && upnParsed.domain !== UPN_DOMAIN.toLowerCase()) {
        messages.push(`User Principal Name: domain must be ${UPN_DOMAIN}`);
      }
    }
  }

  // Mail format & domain
  if (mail) {
    if (!isValidEmailFormat(mail)) {
      messages.push('Mail: invalid format');
    } else {
      const mailParsed = parseEmail(mail);
      const allowed = MAIL_DOMAINS.map((d) => d.toLowerCase());
      if (mailParsed && !allowed.includes(mailParsed.domain)) {
        messages.push(
          `Mail: domain must be one of ${MAIL_DOMAINS.join(', ')}`
        );
      }
    }
  }

  // UPN / Mail local part match
  if (upn && mail) {
    const upnParsed = parseEmail(upn);
    const mailParsed = parseEmail(mail);
    if (
      upnParsed &&
      mailParsed &&
      upnParsed.local.toLowerCase() !== mailParsed.local.toLowerCase()
    ) {
      messages.push('User Principal Name and Mail local part do not match');
    }
  }

  return messages;
}

/**
 * Returns duplicate-error messages for a row (which unique columns have duplicates).
 */
function getRowDuplicateMessages(
  rows: unknown[][],
  rowIndex: number
): string[] {
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
 * Validates the UpdateUsers table data rows.
 * Same rules as CreateUsers verify, plus Object ID must be a valid UUID.
 */
export function verifyUpdateUsers(rows: unknown[][]): VerifyUsersResult {
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
 */
export function getUpdateRowProblemDescription(
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
