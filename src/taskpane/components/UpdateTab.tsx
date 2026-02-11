/**
 * Update Users tab: Initialize Table, Load Data, Verify, and Update actions.
 */

import { useState } from 'react';
import { useAuth } from '@/auth/AuthContext.tsx';
import { loadUserByUpn } from '../loadUsers';
import {
  verifyUpdateUsers,
  NO_INPUT_TABLE_UPDATE_VERIFY_RESULT,
} from '../verifyUpdateUsers';
import { updateUser } from '../updateUser';
import {
  ensureInputAndGetUpdateUsersTable,
  applyUpdateVerifyResultToSheet,
  hasUpdateTableData,
  writeLoadedDataToTable,
  markLoadErrorOnRow,
} from '../excelUpdateUsersTable';
import { initSheetTemplate } from '../initSheetTemplate';
import type { VerifyUsersResult } from '../verifyCore';
import { useAppStyles } from '../App.styles';
import { ActionCard } from './ActionCard';
import { VerifyResultPanel } from './VerifyResultPanel';

const USERS_TABLE_HEADERS = [
  'User Principal Name',
  'Mail',
  'BMS ID',
  'Local HR ID',
  'SN Ticket ID',
  'First Name',
  'Last Name',
  'Display Name',
  'Country',
  'City',
  'Job Title',
  'Office Location',
  'Street Address',
  'State',
  'Postal Code',
  'Business Phone',
  'Mobile Phone',
  'Company Name',
  'Department',
] as const;

const UPDATE_TABLE_HEADERS = ['Object ID', ...USERS_TABLE_HEADERS] as const;

const UPDATE_TABLE_DESCRIPTION =
  'This table is used to update existing users in Entra ID.\n' +
  'Fill in the User Principal Name column, then use Load Data to fetch current values from Entra ID.\n' +
  'Modify the values as needed and run Update. The following conditions apply:\n\n' +
  '\u2022 Required fields: Object ID, User Principal Name, Mail, First Name, Last Name, Display Name, Country, City\n' +
  '\u2022 Object ID must be a valid UUID\n' +
  '\u2022 BMS ID or Local HR ID must be filled\n' +
  '\u2022 BMS ID: digits only, no leading zero\n' +
  '\u2022 UPN domain: majorel.com\n' +
  '\u2022 Mail domain: majorel.com or mj.teleperformance.com\n' +
  '\u2022 UPN and Mail local part must match\n' +
  '\u2022 No duplicates in: Object ID, User Principal Name, Mail, BMS ID, Local HR ID\n' +
  '\u2022 Max. 100 rows';

const UPDATE_COL_WIDTHS: [string, number][] = [
  ['A:A', 210], ['B:B', 150], ['C:C', 200], ['D:D', 70],
  ['E:E', 70], ['F:F', 90], ['G:G', 90], ['H:H', 90],
  ['I:I', 130], ['J:J', 100], ['K:K', 100], ['L:L', 100],
  ['M:M', 100], ['N:N', 100], ['O:O', 100], ['P:P', 100],
  ['Q:Q', 100], ['R:R', 100], ['S:S', 100], ['T:T', 100],
];

export function UpdateTab() {
  const classes = useAppStyles();
  const { getAccessToken } = useAuth();
  const [verifyResult, setVerifyResult] = useState<VerifyUsersResult | null>(null);

  const handleRunInit = async () => {
    try {
      await Excel.run(async (context) => {
        await initSheetTemplate(context, {
          sheetName: 'Update',
          tableName: 'UpdateUsers',
          title: 'Update',
          description: UPDATE_TABLE_DESCRIPTION,
          headers: UPDATE_TABLE_HEADERS,
          lastColumnLetter: 'T',
          columnWidths: UPDATE_COL_WIDTHS,
          descriptionRowHeight: 220,
        });
      });
    } catch (err) {
      console.error('Init Update Template failed:', err);
    }
  };

  const handleRunLoadData = async () => {
    try {
      const accessToken = await getAccessToken(['User.Read.All']);
      await Excel.run(async (context) => {
        const tableData = await ensureInputAndGetUpdateUsersTable(context);
        if (!hasUpdateTableData(tableData)) {
          return;
        }
        const { rows, dataBodyRange } = tableData;

        // Clear all existing notes and fill colors before loading
        dataBodyRange.format.fill.clear();
        const worksheet = dataBodyRange.worksheet;
        const notes = worksheet.notes;
        notes.load('items');
        await context.sync();
        for (const note of notes.items) {
          note.delete();
        }
        await context.sync();

        // Iterate over rows and load user data by UPN (column index 1)
        for (let i = 0; i < rows.length; i++) {
          const upn = String(rows[i]?.[1] ?? '').trim();
          if (!upn) {
            await markLoadErrorOnRow(
              context,
              dataBodyRange,
              i,
              'User Principal Name is empty'
            );
            continue;
          }

          const result = await loadUserByUpn(upn, accessToken);
          if (result.success) {
            await writeLoadedDataToTable(
              context,
              dataBodyRange,
              i,
              result.values
            );
          } else {
            await markLoadErrorOnRow(
              context,
              dataBodyRange,
              i,
              result.error
            );
          }
        }
      });
    } catch (err) {
      console.error('Load Data failed:', err);
    }
  };

  const handleRunVerify = async () => {
    try {
      const result = await Excel.run(async (context) => {
        const tableData = await ensureInputAndGetUpdateUsersTable(context);
        if (!hasUpdateTableData(tableData)) {
          return NO_INPUT_TABLE_UPDATE_VERIFY_RESULT;
        }
        const { rows, dataBodyRange } = tableData;
        const verifyResult = verifyUpdateUsers(rows);
        await applyUpdateVerifyResultToSheet(context, dataBodyRange, rows, verifyResult);
        return verifyResult;
      });

      setVerifyResult(result);
    } catch {
      setVerifyResult({
        success: false,
        totalRows: 0,
        okCount: 0,
        problemCount: 0,
        problemRowIndices: [],
      });
    }
  };

  const handleRunUpdate = async () => {
    try {
      const accessToken = await getAccessToken(['User.ReadWrite.All']);
      const runResult = await Excel.run(async (context) => {
        const tableData = await ensureInputAndGetUpdateUsersTable(context);
        if (!hasUpdateTableData(tableData)) {
          return { verifyResult: NO_INPUT_TABLE_UPDATE_VERIFY_RESULT, updated: false };
        }
        const { rows, dataBodyRange } = tableData;
        const verifyResultData = verifyUpdateUsers(rows);
        await applyUpdateVerifyResultToSheet(
          context,
          dataBodyRange,
          rows,
          verifyResultData
        );

        if (!verifyResultData.success || verifyResultData.noInputTable) {
          return { verifyResult: verifyResultData, updated: false };
        }

        const worksheets = context.workbook.worksheets;
        const sheetName =
          'Update_' +
          new Date()
            .toISOString()
            .slice(0, 19)
            .replace('T', '_')
            .replace(/:/g, '-');
        const newSheet = worksheets.add(sheetName);
        newSheet.activate();

        const LOG_HEADERS = [
          'Timestamp',
          'Object ID',
          'User Principal Name',
          'Display Name',
          'Status',
          'Error',
        ] as const;

        const headerRange = newSheet.getRange('A1:F1');
        headerRange.values = [LOG_HEADERS as unknown as string[]];
        const logTable = newSheet.tables.add('A1:F1', true);

        newSheet.getRange('A:A').format.columnWidth = 170;
        newSheet.getRange('B:B').format.columnWidth = 210;
        newSheet.getRange('C:C').format.columnWidth = 170;
        newSheet.getRange('D:D').format.columnWidth = 170;
        newSheet.getRange('E:E').format.columnWidth = 100;
        newSheet.getRange('F:F').format.columnWidth = 300;

        const logRows: (string | number)[][] = [];
        for (const row of rows) {
          const objectId = String(row[0] ?? '').trim();
          const upn = String(row[1] ?? '').trim();
          const displayName = String(row[8] ?? '').trim();
          const result = await updateUser(row, accessToken);
          logRows.push([
            new Date().toISOString(),
            objectId,
            upn,
            displayName,
            result.status,
            result.error ?? '',
          ]);
        }

        if (logRows.length > 0) {
          logTable.rows.add(undefined, logRows);
        }

        await context.sync();
        return { verifyResult: verifyResultData, updated: true };
      });

      setVerifyResult(runResult.verifyResult);
    } catch (err) {
      console.error('Update failed:', err);
    }
  };

  return (
    <div className={classes.panel} key="update">
      <ActionCard
        title="Initialize Table"
        description="Initializes the Update table for the user data."
        buttonLabel="Initialize"
        onAction={handleRunInit}
      />
      <ActionCard
        title="Load Data"
        description="Loads current user data from Entra ID by User Principal Name."
        buttonLabel="Load"
        onAction={handleRunLoadData}
      />
      <ActionCard
        title="Verify Data"
        description="Verifies the data in the Update table."
        buttonLabel="Verify"
        onAction={handleRunVerify}
      >
        <VerifyResultPanel result={verifyResult} noDataLabel="No Update Data" />
      </ActionCard>
      <ActionCard
        title="Update"
        description="Updates existing user accounts in Entra ID."
        buttonLabel="Update"
        onAction={handleRunUpdate}
      />
    </div>
  );
}
