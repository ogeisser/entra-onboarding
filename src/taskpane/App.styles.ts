import { makeStyles, shorthands } from '@fluentui/react-components';
import { tokens } from '@fluentui/react-theme';

export const useAppStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    minHeight: '100vh',
    boxSizing: 'border-box',
    ...shorthands.padding('16px'),
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    flexShrink: 0,
    ...shorthands.margin(0, 0, '16px', 0),
    ...shorthands.padding(0, 0, '12px', 0),
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
  },
  tabs: {
    flexShrink: 0,
    ...shorthands.margin(0, 0, '16px', 0),
  },
  content: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('20px'),
    ...shorthands.padding('8px', 0, 0, 0),
  },
  panel: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('20px'),
    '@media (prefers-reduced-motion: reduce)': {
      animation: 'none',
    },
  },
  verifyResult: {
    ...shorthands.margin('12px', 0, 0, 0),
    ...shorthands.padding('12px'),
    minHeight: '2.5em',
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
  },
  verifyMessage: {
    ...shorthands.margin(0),
    color: tokens.colorNeutralForeground2,
  },
  verifyStatusMessage: {
    ...shorthands.margin(0),
    /* inherit color from parent verifyStatusLine */
  },
  verifyMessageOk: {
    color: tokens.colorPaletteGreenForeground1,
  },
  verifyMessageErrors: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: 600,
  },
  verifyMessageWarning: {
    color: tokens.colorPaletteMarigoldForeground1,
    fontWeight: 600,
  },
  verifyStats: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('4px'),
  },
  verifyStatusLine: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('8px'),
    fontWeight: 600,
  },
  verifyStatusIcon: {
    width: '28px',
    height: '28px',
    flexShrink: 0,
  },
  verifyTable: {
    ...shorthands.margin('12px', 0, 0, 0),
    width: '100%',
    minWidth: 0,
  },
  verifyTableRowTotal: {
    fontWeight: 700,
  },
  verifyTableRowOk: {
    color: tokens.colorPaletteGreenForeground1,
    fontWeight: 700,
  },
  verifyTableRowErrors: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: 700,
  },
  verifyTableIcon: {
    display: 'block',
    flexShrink: 0,
  },
});
