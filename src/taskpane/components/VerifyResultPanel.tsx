/**
 * Shared Verify Result display panel.
 * Shows verification status (ok/errors/warning) and a summary table.
 */

import {
  Body1,
  mergeClasses,
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
} from '@fluentui/react-components';
import {
  CheckmarkCircle20Regular,
  CheckmarkCircle24Filled,
  ErrorCircle20Regular,
  ErrorCircle24Filled,
  Warning24Filled,
} from '@fluentui/react-icons';
import type { VerifyUsersResult } from '../verifyCore';
import { useAppStyles } from '../App.styles';

interface VerifyResultPanelProps {
  result: VerifyUsersResult | null;
  /** Label shown when noInputTable is true, e.g. "No Create Data" */
  noDataLabel: string;
}

export function VerifyResultPanel({ result, noDataLabel }: VerifyResultPanelProps) {
  const classes = useAppStyles();

  if (result === null) {
    return (
      <div className={classes.verifyResult} aria-live="polite">
        <Body1 className={classes.verifyMessage}>
          Run Verify to check the table.
        </Body1>
      </div>
    );
  }

  return (
    <div className={classes.verifyResult} aria-live="polite">
      <div className={classes.verifyStats}>
        <div
          className={mergeClasses(
            classes.verifyStatusLine,
            result.noInputTable
              ? classes.verifyMessageWarning
              : result.success
                ? classes.verifyMessageOk
                : classes.verifyMessageErrors,
          )}
          role="status"
          aria-label={
            result.noInputTable
              ? noDataLabel
              : result.success
                ? 'OK – Verification passed.'
                : 'Verification completed with problems.'
          }
        >
          {result.noInputTable ? (
            <Warning24Filled className={classes.verifyStatusIcon} />
          ) : result.success ? (
            <CheckmarkCircle24Filled className={classes.verifyStatusIcon} />
          ) : (
            <ErrorCircle24Filled className={classes.verifyStatusIcon} />
          )}
          <Body1 className={classes.verifyStatusMessage}>
            {result.noInputTable
              ? noDataLabel
              : result.success
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
              <TableCell>{result.okCount}</TableCell>
            </TableRow>
            <TableRow className={classes.verifyTableRowErrors}>
              <TableCell>
                <ErrorCircle20Regular
                  className={classes.verifyTableIcon}
                  aria-label="Errors"
                />
              </TableCell>
              <TableCell>{result.problemCount}</TableCell>
            </TableRow>
            <TableRow className={classes.verifyTableRowTotal}>
              <TableCell>Total</TableCell>
              <TableCell>{result.totalRows}</TableCell>
            </TableRow>
          </TableBody>
        </Table>
      </div>
    </div>
  );
}
