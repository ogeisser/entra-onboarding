/**
 * Validation for the "CreateUsers" table.
 * Uses the shared verifyCore engine with Create-specific column configuration.
 */

import {
  createVerifier,
  noInputTableResult,
  type VerifyUsersResult,
} from './verifyCore';

// Re-export shared types and constants so existing imports keep working.
export type { VerifyUsersResult } from './verifyCore';
export { UPN_DOMAIN, MAIL_DOMAINS, MAX_DATA_ROWS } from './verifyCore';

/** Column indices matching USERS_TABLE_HEADERS in App.tsx. */
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

const createVerifierInstance = createVerifier({
  columns: {
    userPrincipalName: COL.userPrincipalName,
    mail: COL.mail,
    bmsId: COL.bmsId,
    localHrId: COL.localHrId,
  },
  requiredColumns: [
    { index: COL.userPrincipalName, name: 'User Principal Name' },
    { index: COL.mail, name: 'Mail' },
    { index: COL.firstName, name: 'First Name' },
    { index: COL.lastName, name: 'Last Name' },
    { index: COL.displayName, name: 'Display Name' },
    { index: COL.country, name: 'Country' },
    { index: COL.city, name: 'City' },
  ],
  uniqueColumns: [
    { index: COL.userPrincipalName, name: 'User Principal Name' },
    { index: COL.mail, name: 'Mail' },
    { index: COL.bmsId, name: 'BMS ID' },
    { index: COL.localHrId, name: 'Local HR ID' },
  ],
});

/** Result to return when the CreateUsers table was not found (no create data). */
export const NO_INPUT_TABLE_VERIFY_RESULT: VerifyUsersResult = noInputTableResult();

/** Validates the CreateUsers table data rows. */
export const verifyUsers = createVerifierInstance.verify;

/** Returns a human-readable problem description for a row, or empty string. */
export const getRowProblemDescription = createVerifierInstance.getRowProblemDescription;
