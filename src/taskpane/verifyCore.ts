/**
 * Shared verification engine for user tables (Create and Update).
 * Provides a factory function that creates a verifier based on column configuration.
 */

export const UPN_DOMAIN = 'majorel.com';
export const MAIL_DOMAINS = ['majorel.com', 'mj.teleperformance.com'] as const;
export const MAX_DATA_ROWS = 100;

export interface VerifyUsersResult {
  success: boolean;
  totalRows: number;
  okCount: number;
  problemCount: number;
  /** 0-based row indices (in the rows array) that have validation errors */
  problemRowIndices: number[];
  /** true when the input table was not found (no data) */
  noInputTable?: boolean;
}

/** Creates a NO_INPUT_TABLE result. */
export function noInputTableResult(): VerifyUsersResult {
  return {
    success: false,
    totalRows: 0,
    okCount: 0,
    problemCount: 0,
    problemRowIndices: [],
    noInputTable: true,
  };
}

// ─── Shared helpers ──────────────────────────────────────────────────

export function cellValue(row: unknown[], index: number): string {
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

/** BMS ID: digits only, no leading zero (e.g. "0" or "123" ok, "01" / "007" invalid). */
function isValidBmsId(value: string): boolean {
  if (value === '') return true;
  return /^(0|[1-9]\d*)$/.test(value);
}

// ─── Verifier factory ────────────────────────────────────────────────

/** Column indices that the shared validation logic needs. */
export interface VerifierColumnConfig {
  userPrincipalName: number;
  mail: number;
  bmsId: number;
  localHrId: number;
}

export interface VerifierConfig {
  /** Column indices for the shared field checks (UPN, Mail, BMS ID, Local HR ID). */
  columns: VerifierColumnConfig;
  /** Columns that must not be empty. */
  requiredColumns: { index: number; name: string }[];
  /** Columns whose values must be unique across all rows. */
  uniqueColumns: { index: number; name: string }[];
  /** Optional extra validators run per row (return error messages). */
  extraValidators?: ((row: unknown[]) => string[])[];
}

function getRowValidationErrorMessages(
  row: unknown[],
  config: VerifierConfig
): string[] {
  const messages: string[] = [];
  const { columns, requiredColumns, extraValidators } = config;

  // Required fields
  for (const { index, name } of requiredColumns) {
    if (cellValue(row, index) === '') {
      messages.push(`Required field '${name}' is empty`);
    }
  }

  // Extra validators (e.g. UUID check for Object ID)
  if (extraValidators) {
    for (const validator of extraValidators) {
      messages.push(...validator(row));
    }
  }

  // BMS ID or Local HR ID must be present
  const bmsId = cellValue(row, columns.bmsId);
  const localHrId = cellValue(row, columns.localHrId);
  if (bmsId === '' && localHrId === '') {
    messages.push('BMS ID and Local HR ID are both empty');
  }
  if (bmsId !== '' && !isValidBmsId(bmsId)) {
    messages.push('BMS ID must be a number (digits only, no leading zero)');
  }

  // UPN format & domain
  const upn = cellValue(row, columns.userPrincipalName);
  const mail = cellValue(row, columns.mail);

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
        messages.push(`Mail: domain must be one of ${MAIL_DOMAINS.join(', ')}`);
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

function getRowDuplicateMessages(
  rows: unknown[][],
  rowIndex: number,
  uniqueColumns: { index: number; name: string }[]
): string[] {
  const messages: string[] = [];
  const row = rows[rowIndex] ?? [];

  for (const { index, name } of uniqueColumns) {
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

function findRowsWithDuplicateUniqueValues(
  rows: unknown[][],
  uniqueColumns: { index: number; name: string }[]
): Set<number> {
  const duplicateRowIndices = new Set<number>();

  for (const { index } of uniqueColumns) {
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

export interface Verifier {
  /** Validates rows and returns the result summary. */
  verify(rows: unknown[][]): VerifyUsersResult;
  /** Returns a human-readable problem description for a single row, or empty string. */
  getRowProblemDescription(rows: unknown[][], rowIndex: number): string;
}

/**
 * Creates a verifier with the given column configuration.
 * Used by both CreateUsers and UpdateUsers validation.
 */
export function createVerifier(config: VerifierConfig): Verifier {
  const PROBLEM_MESSAGE_SEPARATOR = '\n';

  function verify(rows: unknown[][]): VerifyUsersResult {
    const totalRows = rows.length;
    const problemRowIndicesSet = new Set<number>();

    for (let i = 0; i < rows.length; i++) {
      if (getRowValidationErrorMessages(rows[i], config).length > 0) {
        problemRowIndicesSet.add(i);
      }
    }

    const duplicateRows = findRowsWithDuplicateUniqueValues(rows, config.uniqueColumns);
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

  function getRowProblemDescription(
    rows: unknown[][],
    rowIndex: number
  ): string {
    const row = rows[rowIndex];
    if (!row) return '';
    const validation = getRowValidationErrorMessages(row, config);
    const duplicates = getRowDuplicateMessages(rows, rowIndex, config.uniqueColumns);
    const all = [...validation, ...duplicates];
    return all.length === 0 ? '' : all.join(PROBLEM_MESSAGE_SEPARATOR);
  }

  return { verify, getRowProblemDescription };
}
