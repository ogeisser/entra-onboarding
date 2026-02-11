/**
 * Validation for the "UpdateUsers" table.
 * Uses the shared verifyCore engine with Update-specific column configuration.
 * Object ID at index 0, then same columns as Create shifted by 1.
 */

import {
  cellValue,
  createVerifier,
  noInputTableResult,
  type VerifyUsersResult,
} from './verifyCore';

// Re-export for convenience.
export type { VerifyUsersResult } from './verifyCore';

const UUID_REGEX =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

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

const updateVerifierInstance = createVerifier({
  columns: {
    userPrincipalName: COL.userPrincipalName,
    mail: COL.mail,
    bmsId: COL.bmsId,
    localHrId: COL.localHrId,
  },
  requiredColumns: [
    { index: COL.objectId, name: 'Object ID' },
    { index: COL.userPrincipalName, name: 'User Principal Name' },
    { index: COL.mail, name: 'Mail' },
    { index: COL.firstName, name: 'First Name' },
    { index: COL.lastName, name: 'Last Name' },
    { index: COL.displayName, name: 'Display Name' },
    { index: COL.country, name: 'Country' },
    { index: COL.city, name: 'City' },
  ],
  uniqueColumns: [
    { index: COL.objectId, name: 'Object ID' },
    { index: COL.userPrincipalName, name: 'User Principal Name' },
    { index: COL.mail, name: 'Mail' },
    { index: COL.bmsId, name: 'BMS ID' },
    { index: COL.localHrId, name: 'Local HR ID' },
  ],
  extraValidators: [
    (row: unknown[]) => {
      const objectId = cellValue(row, COL.objectId);
      if (objectId !== '' && !UUID_REGEX.test(objectId)) {
        return ['Object ID must be a valid UUID'];
      }
      return [];
    },
  ],
});

/** Result to return when the UpdateUsers table was not found (no update data). */
export const NO_INPUT_TABLE_UPDATE_VERIFY_RESULT: VerifyUsersResult = noInputTableResult();

/** Validates the UpdateUsers table data rows. */
export const verifyUpdateUsers = updateVerifierInstance.verify;

/** Returns a human-readable problem description for a row, or empty string. */
export const getUpdateRowProblemDescription = updateVerifierInstance.getRowProblemDescription;
