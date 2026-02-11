/**
 * Create Users tab: Initialize Table, Verify, and Create actions.
 */

import { useState } from 'react';
import { useAuth } from '@/auth/AuthContext.tsx';
import { createUser } from '../createUser';
import {
  verifyUsers,
  NO_INPUT_TABLE_VERIFY_RESULT,
  type VerifyUsersResult,
} from '../verifyUsers';
import {
  ensureInputAndGetUsersTable,
  applyVerifyResultToSheet,
  hasTableData,
} from '../excelUsersTable';
import { initSheetTemplate } from '../initSheetTemplate';
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

const CREATE_TABLE_DESCRIPTION =
  'This table is used to create new users in Entra ID.\n' +
  'Data is validated by Verify before creation. The following conditions apply:\n\n' +
  '\u2022 Required fields: User Principal Name, Mail, First Name, Last Name, Display Name, Country, City\n' +
  '\u2022 BMS ID or Local HR ID must be filled\n' +
  '\u2022 BMS ID: digits only, no leading zero\n' +
  '\u2022 UPN domain: majorel.com\n' +
  '\u2022 Mail domain: majorel.com or mj.teleperformance.com\n' +
  '\u2022 UPN and Mail local part must match\n' +
  '\u2022 No duplicates in: User Principal Name, Mail, BMS ID, Local HR ID\n' +
  '\u2022 Max. 100 rows';

const CREATE_COL_WIDTHS: [string, number][] = [
  ['A:A', 150], ['B:B', 200], ['C:C', 70], ['D:D', 70],
  ['E:E', 90], ['F:F', 90], ['G:G', 90], ['H:H', 130],
  ['I:I', 100], ['J:J', 100], ['K:K', 100], ['L:L', 100],
  ['M:M', 100], ['N:N', 100], ['O:O', 100], ['P:P', 100],
  ['Q:Q', 100], ['R:R', 100], ['S:S', 100],
];

export function CreateTab() {
  const classes = useAppStyles();
  const { getAccessToken } = useAuth();
  const [verifyResult, setVerifyResult] = useState<VerifyUsersResult | null>(null);

  const handleRunInit = async () => {
    try {
      await Excel.run(async (context) => {
        await initSheetTemplate(context, {
          sheetName: 'Create',
          tableName: 'CreateUsers',
          title: 'Create',
          description: CREATE_TABLE_DESCRIPTION,
          headers: USERS_TABLE_HEADERS,
          lastColumnLetter: 'S',
          columnWidths: CREATE_COL_WIDTHS,
          descriptionRowHeight: 185,
        });
      });
    } catch (err) {
      console.error('Init Template failed:', err);
    }
  };

  const handleRunVerify = async () => {
    try {
      const result = await Excel.run(async (context) => {
        const tableData = await ensureInputAndGetUsersTable(context);
        if (!hasTableData(tableData)) {
          return NO_INPUT_TABLE_VERIFY_RESULT;
        }
        const { rows, dataBodyRange } = tableData;
        const verifyResult = verifyUsers(rows);
        await applyVerifyResultToSheet(context, dataBodyRange, rows, verifyResult);
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

  const handleRunCreate = async () => {
    try {
      const accessToken = await getAccessToken(['User.ReadWrite.All']);
      const runResult = await Excel.run(async (context) => {
        const tableData = await ensureInputAndGetUsersTable(context);
        if (!hasTableData(tableData)) {
          return { verifyResult: NO_INPUT_TABLE_VERIFY_RESULT, created: false };
        }
        const { rows, dataBodyRange } = tableData;
        const verifyResultData = verifyUsers(rows);
        await applyVerifyResultToSheet(
          context,
          dataBodyRange,
          rows,
          verifyResultData
        );

        if (!verifyResultData.success || verifyResultData.noInputTable) {
          return { verifyResult: verifyResultData, created: false };
        }

        const worksheets = context.workbook.worksheets;
        const sheetName =
          'Create_' +
          new Date()
            .toISOString()
            .slice(0, 19)
            .replace('T', '_')
            .replace(/:/g, '-');
        const newSheet = worksheets.add(sheetName);
        newSheet.activate();

        const LOG_HEADERS = [
          'Timestamp',
          'User Principal Name',
          'Display Name',
          'Object Id',
          'Generated Password',
          'Status',
          'Error',
        ] as const;

        const headerRange = newSheet.getRange('A1:G1');
        headerRange.values = [LOG_HEADERS as unknown as string[]];
        const logTable = newSheet.tables.add('A1:G1', true);

        newSheet.getRange('A:C').format.columnWidth = 170;
        newSheet.getRange('D:D').format.columnWidth = 210;
        newSheet.getRange('E:F').format.columnWidth = 170;
        newSheet.getRange('G:G').format.columnWidth = 210;

        const logRows: (string | number)[][] = [];
        for (const row of rows) {
          const upn = String(row[0] ?? '').trim();
          const displayName = String(row[7] ?? '').trim();
          const result = await createUser(row, accessToken);
          logRows.push([
            new Date().toISOString(),
            upn,
            displayName,
            result.objectId ?? '',
            result.generatedPassword ?? '',
            result.status,
            result.error ?? '',
          ]);
        }

        if (logRows.length > 0) {
          logTable.rows.add(undefined, logRows);
        }

        await context.sync();
        return { verifyResult: verifyResultData, created: true };
      });

      setVerifyResult(runResult.verifyResult);
    } catch (err) {
      console.error('Create failed:', err);
    }
  };

  return (
    <div className={classes.panel} key="create">
      <ActionCard
        title="Initialize Table"
        description="Initializes the Create table for the user data."
        buttonLabel="Initialize"
        onAction={handleRunInit}
      />
      <ActionCard
        title="Verify Data"
        description="Verifies the data in the Create table."
        buttonLabel="Verify"
        onAction={handleRunVerify}
      >
        <VerifyResultPanel result={verifyResult} noDataLabel="No Create Data" />
      </ActionCard>
      <ActionCard
        title="Create"
        description="Creates new user accounts in Entra ID."
        buttonLabel="Create"
        onAction={handleRunCreate}
      />
    </div>
  );
}
