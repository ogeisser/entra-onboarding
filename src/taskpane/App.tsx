import { useCallback, useEffect, useState, type ReactNode } from 'react';
import type { SelectTabData, TabValue } from '@fluentui/react-components';
import {
  Body1,
  Button,
  Card,
  CardFooter,
  CardHeader,
  mergeClasses,
  Tab,
  TabList,
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
  Title3,
} from '@fluentui/react-components';
import { ProfileMenu } from './components/ProfileMenu';
import { useAuth } from '@/auth/AuthContext.tsx';
import { createUser } from './createUser';
import {
  verifyUsers,
  NO_INPUT_TABLE_VERIFY_RESULT,
  type VerifyUsersResult,
} from './verifyUsers';
import {
  ensureInputAndGetUsersTable,
  applyVerifyResultToSheet,
  hasTableData,
} from './excelUsersTable';
import { loadUserByUpn } from './loadUsers';
import {
  verifyUpdateUsers,
  NO_INPUT_TABLE_UPDATE_VERIFY_RESULT,
} from './verifyUpdateUsers';
import { updateUser } from './updateUser';
import {
  ensureInputAndGetUpdateUsersTable,
  applyUpdateVerifyResultToSheet,
  hasUpdateTableData,
  writeLoadedDataToTable,
  markLoadErrorOnRow,
} from './excelUpdateUsersTable';
import { useAppStyles } from './App.styles';
import {
  CheckmarkCircle20Regular,
  CheckmarkCircle24Filled,
  ErrorCircle20Regular,
  ErrorCircle24Filled,
  Warning24Filled,
} from '@fluentui/react-icons';

/** Vorübergehend: formatierte Fehlerdetails für Auth-Debugging (inkl. NAA/Bridge-Fehler) */
function authErrorDescription(error: Error): ReactNode {
  const e = error as Error & Record<string, unknown>;
  const noMessage = !error.message || error.message === '(no message)';
  const isUserCanceled = e.errorCode === 'user_canceled' || (error.message && error.message.includes('user_canceled'));
  const isServerErrorNoDetails = e.name === 'ServerError' && noMessage && !e.errorCode && !e.errorMessage;

  const parts: ReactNode[] = [];
  if (isUserCanceled) {
    parts.push(<span key="cancel">Anmeldung abgebrochen. Sie können jederzeit erneut auf «Sign in» tippen.</span>);
  } else if (isServerErrorNoDetails) {
    parts.push(
      <span key="hint">
        Die Office-Anmeldung (NAA-Bridge) ist fehlgeschlagen, ohne genaue Fehlermeldung. Das passiert oft in Excel im Browser oder bei temporären Verbindungsproblemen.
        <br /><br />
        <strong>Was Sie tun können:</strong>
        <br />• Erneut auf «Sign in» tippen – es wird dann ein Anmelde-Popup geöffnet.
        <br />• Excel schließen und neu starten, dann erneut anmelden.
        <br />• Falls Sie Excel im Browser nutzen: Desktop-Excel ausprobieren (oder umgekehrt).
      </span>
    );
  } else {
    parts.push(<span key="msg">{error.message || '(keine Meldung)'}</span>);
  }
  const hideDetails = isUserCanceled || isServerErrorNoDetails;
  if (e.name && !hideDetails) parts.push(<><br /><strong>Name:</strong> {String(e.name)}</>);
  if (e.errorCode !== undefined && e.errorCode !== '' && !hideDetails) parts.push(<><br /><strong>ErrorCode:</strong> {String(e.errorCode)}</>);
  if (e.errorMessage !== undefined && e.errorMessage !== '' && !hideDetails) parts.push(<><br /><strong>ErrorMessage:</strong> {String(e.errorMessage)}</>);
  if (e.subError && !hideDetails) parts.push(<><br /><strong>SubError:</strong> {String(e.subError)}</>);
  if (e.status != null && !hideDetails) parts.push(<><br /><strong>Status:</strong> {String(e.status)}</>);
  if (e.code && !hideDetails) parts.push(<><br /><strong>Code:</strong> {String(e.code)}</>);
  if (e.description && !hideDetails) parts.push(<><br /><strong>Description:</strong> {String(e.description)}</>);
  if (error.stack && !hideDetails) parts.push(<><br /><strong>Stack:</strong><pre style={{ fontSize: '0.75rem', whiteSpace: 'pre-wrap' as const, marginTop: 4 }}>{error.stack}</pre></>);
  return <>{parts}</>;
}

function App() {
  const classes = useAppStyles();
  const {
    getAccessToken,
    isNaaSupported,
    naaAvailable,
    naaInitError,
    user,
    loginInProgress,
    login,
    logout,
    error: authError,
    clearError,
  } = useAuth();
  const [selectedTab, setSelectedTab] = useState<TabValue>('create');
  const [verifyResult, setVerifyResult] = useState<VerifyUsersResult | null>(null);
  const [updateVerifyResult, setUpdateVerifyResult] = useState<VerifyUsersResult | null>(null);
  const logDebugInfo = useCallback(async () => {
    try {
      const officeDiagnostics =
        typeof Office !== 'undefined' && Office.context?.diagnostics
          ? (Office.context.diagnostics as unknown as Record<string, unknown>)
          : null;
      const hasGetAuthContext =
        typeof (Office as unknown as { auth?: { getAuthContext?: unknown } })?.auth?.getAuthContext ===
        'function';

      let nestedAppAuthRequirementMet: boolean | null = null;
      try {
        if (typeof Office !== 'undefined' && typeof Office.context?.requirements?.isSetSupported === 'function') {
          nestedAppAuthRequirementMet = Office.context.requirements.isSetSupported('NestedAppAuth', '1.1');
        }
      } catch {
        // keep null
      }

      let authContext: Record<string, string> | null = null;
      let authContextError: string | undefined;

      if (hasGetAuthContext) {
        try {
          const ctx = await Office.auth.getAuthContext();
          authContext = {
            userObjectId: ctx.userObjectId ?? '',
            tenantId: ctx.tenantId ?? '',
            userPrincipalName: ctx.userPrincipalName ?? '',
            authorityType: ctx.authorityType ?? '',
            authorityBaseUrl: ctx.authorityBaseUrl ?? '',
            loginHint: ctx.loginHint ?? '',
          };
        } catch (err) {
          authContextError = err instanceof Error ? err.message : String(err);
        }
      }

      console.log('[Debug Info]', {
        officeDiagnostics,
        hasGetAuthContext,
        authContext,
        authContextError,
        nestedAppAuthRequirementMet,
      });
    } catch (err) {
      console.log('[Debug Info]', {
        officeDiagnostics: null,
        hasGetAuthContext: false,
        authContext: null,
        nestedAppAuthRequirementMet: null,
      });
    }
  }, []);

  useEffect(() => {
    if (user === null) {
      logDebugInfo();
    }
  }, [user, logDebugInfo]);

  const onTabSelect = (_: unknown, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

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

  const handleRunInit = async () => {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;

        // Vorhandenes "Create"-Sheet löschen, falls es existiert
        const existingCreateSheet = worksheets.getItemOrNullObject('Create');
        await context.sync();

        if (!existingCreateSheet.isNullObject) {
          existingCreateSheet.delete();
          await context.sync();
        }

        // "Create"-Sheet neu anlegen
        const createSheet = worksheets.add('Create');
        await context.sync();

        // Gesamtes Sheet auf Datentyp "Text" setzen (Formatcode @)
        const entireSheetRange = createSheet.getRange();
        (entireSheetRange as unknown as { numberFormat: string }).numberFormat = '@';

        // "Create" zum aktiven Sheet machen
        createSheet.activate();
        await context.sync();

        // Zeile 1: Wert in A1 setzen, dann A1:S1 verbinden und formatieren
        const titleCell = createSheet.getRange('A1');
        titleCell.values = [['Create']];
        const titleRow = createSheet.getRange('A1:S1');
        titleRow.merge();
        titleRow.format.font.bold = true;
        titleRow.format.font.size = 16;
        titleRow.format.font.color = '#FFFFFF';
        titleRow.format.fill.color = '#0078D4';
        titleRow.format.horizontalAlignment = Excel.HorizontalAlignment.left;
        titleRow.format.verticalAlignment = Excel.VerticalAlignment.center;
        titleRow.format.rowHeight = 25;

        // Zeile 2: Wert in A2 setzen, dann A2:S2 verbinden und formatieren
        const descCell = createSheet.getRange('A2');
        descCell.values = [[CREATE_TABLE_DESCRIPTION]];
        const descRow = createSheet.getRange('A2:S2');
        descRow.merge();
        descRow.format.wrapText = true;
        descRow.format.verticalAlignment = Excel.VerticalAlignment.top;
        descRow.format.font.size = 12;
        descRow.format.rowHeight = 185;

        // Spaltenbreiten setzen (A–S)
        const colWidths: [string, number][] = [
          ['A:A', 150], ['B:B', 200], ['C:C', 70], ['D:D', 70],
          ['E:E', 90], ['F:F', 90], ['G:G', 90], ['H:H', 130],
          ['I:I', 100], ['J:J', 100], ['K:K', 100], ['L:L', 100],
          ['M:M', 100], ['N:N', 100], ['O:O', 100], ['P:P', 100],
          ['Q:Q', 100], ['R:R', 100], ['S:S', 100],
        ];
        for (const [col, width] of colWidths) {
          createSheet.getRange(col).format.columnWidth = width;
        }

        // Zeile 3: Header schreiben und Tabelle "CreateUsers" anlegen
        const headerRange = createSheet.getRange('A3:S3');
        headerRange.values = [USERS_TABLE_HEADERS as unknown as string[]];

        const tables = createSheet.tables;
        const newTable = tables.add('A3:S3', true);
        newTable.name = 'CreateUsers';

        await context.sync();
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

        // Spaltenbreiten für Protokoll-Sheet setzen (A–G: 170, D: 210)
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

  // ─── Update Tab Handlers ───────────────────────────────────────────

  const handleRunUpdateInit = async () => {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;

        // Vorhandenes "Update"-Sheet löschen, falls es existiert
        const existingUpdateSheet = worksheets.getItemOrNullObject('Update');
        await context.sync();

        if (!existingUpdateSheet.isNullObject) {
          existingUpdateSheet.delete();
          await context.sync();
        }

        // "Update"-Sheet neu anlegen
        const updateSheet = worksheets.add('Update');
        await context.sync();

        // Gesamtes Sheet auf Datentyp "Text" setzen (Formatcode @)
        const entireSheetRange = updateSheet.getRange();
        (entireSheetRange as unknown as { numberFormat: string }).numberFormat = '@';

        // "Update" zum aktiven Sheet machen
        updateSheet.activate();
        await context.sync();

        // Zeile 1: Wert in A1 setzen, dann A1:T1 verbinden und formatieren
        const titleCell = updateSheet.getRange('A1');
        titleCell.values = [['Update']];
        const titleRow = updateSheet.getRange('A1:T1');
        titleRow.merge();
        titleRow.format.font.bold = true;
        titleRow.format.font.size = 16;
        titleRow.format.font.color = '#FFFFFF';
        titleRow.format.fill.color = '#0078D4';
        titleRow.format.horizontalAlignment = Excel.HorizontalAlignment.left;
        titleRow.format.verticalAlignment = Excel.VerticalAlignment.center;
        titleRow.format.rowHeight = 25;

        // Zeile 2: Wert in A2 setzen, dann A2:T2 verbinden und formatieren
        const descCell = updateSheet.getRange('A2');
        descCell.values = [[UPDATE_TABLE_DESCRIPTION]];
        const descRow = updateSheet.getRange('A2:T2');
        descRow.merge();
        descRow.format.wrapText = true;
        descRow.format.verticalAlignment = Excel.VerticalAlignment.top;
        descRow.format.font.size = 12;
        descRow.format.rowHeight = 220;

        // Spaltenbreiten setzen (A–T: Object ID first, then same as Create shifted)
        const colWidths: [string, number][] = [
          ['A:A', 210], ['B:B', 150], ['C:C', 200], ['D:D', 70],
          ['E:E', 70], ['F:F', 90], ['G:G', 90], ['H:H', 90],
          ['I:I', 130], ['J:J', 100], ['K:K', 100], ['L:L', 100],
          ['M:M', 100], ['N:N', 100], ['O:O', 100], ['P:P', 100],
          ['Q:Q', 100], ['R:R', 100], ['S:S', 100], ['T:T', 100],
        ];
        for (const [col, width] of colWidths) {
          updateSheet.getRange(col).format.columnWidth = width;
        }

        // Zeile 3: Header schreiben und Tabelle "UpdateUsers" anlegen
        const headerRange = updateSheet.getRange('A3:T3');
        headerRange.values = [UPDATE_TABLE_HEADERS as unknown as string[]];

        const tables = updateSheet.tables;
        const newTable = tables.add('A3:T3', true);
        newTable.name = 'UpdateUsers';

        await context.sync();
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

  const handleRunUpdateVerify = async () => {
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

      setUpdateVerifyResult(result);
    } catch {
      setUpdateVerifyResult({
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

        // Spaltenbreiten für Protokoll-Sheet setzen
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

      setUpdateVerifyResult(runResult.verifyResult);
    } catch (err) {
      console.error('Update failed:', err);
    }
  };

  if (user === null) {
    return (
      <div className={classes.container}>
        <div className={classes.content}>
          {authError && (
            <Card>
              <CardHeader
                header={<Title3>Sign-in error</Title3>}
                description={<Body1>{authErrorDescription(authError)}</Body1>}
              />
              <CardFooter action={<Button onClick={clearError}>Dismiss</Button>} />
            </Card>
          )}
          <Card>
            <CardHeader
              header={<Title3>Sign in</Title3>}
              description={
                <Body1>
                  {naaAvailable
                    ? 'Sign in using your Excel account.'
                    : !isNaaSupported
                      ? 'Office too old. Please use a supported Office version.'
                      : 'Nested App Authentication could not be initialized. Sign-in is disabled.' + (naaInitError?.message ? ` (${naaInitError.message})` : '')}
                </Body1>
              }
            />
            <CardFooter
              action={
                <Button
                  appearance="primary"
                  onClick={login}
                  disabled={!naaAvailable || loginInProgress}
                >
                  {loginInProgress ? 'Signing in…' : 'Sign in'}
                </Button>
              }
            />
          </Card>
        </div>
      </div>
    );
  }

  return (
    <div className={classes.container}>
      <header className={classes.header}>
        <ProfileMenu user={user} onLogout={logout} />
      </header>
      <div className={classes.tabs}>
        <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
          <Tab value="create">Create Users</Tab>
          <Tab value="update">Update Users</Tab>
        </TabList>
      </div>

      <div className={classes.content}>
        {authError && (
          <Card>
            <CardHeader
              header={<Title3>Sign-in error</Title3>}
              description={<Body1>{authErrorDescription(authError)}</Body1>}
            />
            <CardFooter action={<Button onClick={clearError}>Dismiss</Button>} />
          </Card>
        )}
        {selectedTab === 'create' && (
          <div className={classes.panel} key="create">
            <Card>
              <CardHeader
                header={<Title3>Initialize Table</Title3>}
                description={
                  <Body1>
                    Initializes the Create table for the user data.
                  </Body1>
                }
              />
              <CardFooter
                action={
                  <Button appearance="primary" onClick={handleRunInit}>
                    Initialize
                  </Button>
                }
              />
            </Card>
            <Card>
              <CardHeader
                header={<Title3>Verify Data</Title3>}
                description={
                  <Body1>
                    Verifies the data in the Create table.
                  </Body1>
                }
              />
              <CardFooter
                action={
                  <Button appearance="primary" onClick={handleRunVerify}>
                    Verify
                  </Button>
                }
              />
              <div className={classes.verifyResult} aria-live="polite">
                {verifyResult === null && (
                  <Body1 className={classes.verifyMessage}>
                    Run Verify to check the table.
                  </Body1>
                )}
                {verifyResult !== null && (
                  <div className={classes.verifyStats}>
                    <div
                      className={mergeClasses(
                        classes.verifyStatusLine,
                        verifyResult.noInputTable
                          ? classes.verifyMessageWarning
                          : verifyResult.success
                            ? classes.verifyMessageOk
                            : classes.verifyMessageErrors,
                      )}
                      role="status"
                      aria-label={
                        verifyResult.noInputTable
                          ? 'No create data.'
                          : verifyResult.success
                            ? 'OK – Verification passed.'
                            : 'Verification completed with problems.'
                      }
                    >
                      {verifyResult.noInputTable ? (
                        <Warning24Filled
                          className={classes.verifyStatusIcon}
                        />
                      ) : verifyResult.success ? (
                        <CheckmarkCircle24Filled
                          className={classes.verifyStatusIcon}
                        />
                      ) : (
                        <ErrorCircle24Filled
                          className={classes.verifyStatusIcon}
                        />
                      )}
                      <Body1 className={classes.verifyStatusMessage}>
                        {verifyResult.noInputTable
                          ? 'No Create Data'
                          : verifyResult.success
                            ? 'Verification passed.'
                            : 'Verification completed with problems.'}
                      </Body1>
                    </div>
                    <Table className={classes.verifyTable} size="small">
                      <TableHeader>
                        <TableRow>
                          <TableHeaderCell>Status</TableHeaderCell>
                          <TableHeaderCell>Rows</TableHeaderCell>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        <TableRow className={classes.verifyTableRowOk}>
                          <TableCell>
                            <CheckmarkCircle20Regular
                              className={classes.verifyTableIcon}
                              aria-label="OK"
                            />
                          </TableCell>
                          <TableCell>{verifyResult.okCount}</TableCell>
                        </TableRow>
                        <TableRow className={classes.verifyTableRowErrors}>
                          <TableCell>
                            <ErrorCircle20Regular
                              className={classes.verifyTableIcon}
                              aria-label="Errors"
                            />
                          </TableCell>
                          <TableCell>{verifyResult.problemCount}</TableCell>
                        </TableRow>
                        <TableRow className={classes.verifyTableRowTotal}>
                          <TableCell>Total</TableCell>
                          <TableCell>{verifyResult.totalRows}</TableCell>
                        </TableRow>
                      </TableBody>
                    </Table>
                  </div>
                )}
              </div>
            </Card>
            <Card>
              <CardHeader
                header={<Title3>Create</Title3>}
                description={
                  <Body1>Creates new user accounts in Entra ID.</Body1>
                }
              />
              <CardFooter
                action={
                  <Button appearance="primary" onClick={handleRunCreate}>
                    Create
                  </Button>
                }
              />
            </Card>
          </div>
        )}
        {selectedTab === 'update' && (
          <div className={classes.panel} key="update">
            <Card>
              <CardHeader
                header={<Title3>Initialize Table</Title3>}
                description={
                  <Body1>
                    Initializes the Update table for the user data.
                  </Body1>
                }
              />
              <CardFooter
                action={
                  <Button appearance="primary" onClick={handleRunUpdateInit}>
                    Initialize
                  </Button>
                }
              />
            </Card>
            <Card>
              <CardHeader
                header={<Title3>Load Data</Title3>}
                description={
                  <Body1>
                    Loads current user data from Entra ID by User Principal Name.
                  </Body1>
                }
              />
              <CardFooter
                action={
                  <Button appearance="primary" onClick={handleRunLoadData}>
                    Load
                  </Button>
                }
              />
            </Card>
            <Card>
              <CardHeader
                header={<Title3>Verify Data</Title3>}
                description={
                  <Body1>
                    Verifies the data in the Update table.
                  </Body1>
                }
              />
              <CardFooter
                action={
                  <Button appearance="primary" onClick={handleRunUpdateVerify}>
                    Verify
                  </Button>
                }
              />
              <div className={classes.verifyResult} aria-live="polite">
                {updateVerifyResult === null && (
                  <Body1 className={classes.verifyMessage}>
                    Run Verify to check the table.
                  </Body1>
                )}
                {updateVerifyResult !== null && (
                  <div className={classes.verifyStats}>
                    <div
                      className={mergeClasses(
                        classes.verifyStatusLine,
                        updateVerifyResult.noInputTable
                          ? classes.verifyMessageWarning
                          : updateVerifyResult.success
                            ? classes.verifyMessageOk
                            : classes.verifyMessageErrors,
                      )}
                      role="status"
                      aria-label={
                        updateVerifyResult.noInputTable
                          ? 'No update data.'
                          : updateVerifyResult.success
                            ? 'OK – Verification passed.'
                            : 'Verification completed with problems.'
                      }
                    >
                      {updateVerifyResult.noInputTable ? (
                        <Warning24Filled
                          className={classes.verifyStatusIcon}
                        />
                      ) : updateVerifyResult.success ? (
                        <CheckmarkCircle24Filled
                          className={classes.verifyStatusIcon}
                        />
                      ) : (
                        <ErrorCircle24Filled
                          className={classes.verifyStatusIcon}
                        />
                      )}
                      <Body1 className={classes.verifyStatusMessage}>
                        {updateVerifyResult.noInputTable
                          ? 'No Update Data'
                          : updateVerifyResult.success
                            ? 'Verification passed.'
                            : 'Verification completed with problems.'}
                      </Body1>
                    </div>
                    <Table className={classes.verifyTable} size="small">
                      <TableHeader>
                        <TableRow>
                          <TableHeaderCell>Status</TableHeaderCell>
                          <TableHeaderCell>Rows</TableHeaderCell>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        <TableRow className={classes.verifyTableRowOk}>
                          <TableCell>
                            <CheckmarkCircle20Regular
                              className={classes.verifyTableIcon}
                              aria-label="OK"
                            />
                          </TableCell>
                          <TableCell>{updateVerifyResult.okCount}</TableCell>
                        </TableRow>
                        <TableRow className={classes.verifyTableRowErrors}>
                          <TableCell>
                            <ErrorCircle20Regular
                              className={classes.verifyTableIcon}
                              aria-label="Errors"
                            />
                          </TableCell>
                          <TableCell>{updateVerifyResult.problemCount}</TableCell>
                        </TableRow>
                        <TableRow className={classes.verifyTableRowTotal}>
                          <TableCell>Total</TableCell>
                          <TableCell>{updateVerifyResult.totalRows}</TableCell>
                        </TableRow>
                      </TableBody>
                    </Table>
                  </div>
                )}
              </div>
            </Card>
            <Card>
              <CardHeader
                header={<Title3>Update</Title3>}
                description={
                  <Body1>Updates existing user accounts in Entra ID.</Body1>
                }
              />
              <CardFooter
                action={
                  <Button appearance="primary" onClick={handleRunUpdate}>
                    Update
                  </Button>
                }
              />
            </Card>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
