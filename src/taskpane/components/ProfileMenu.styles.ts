import { makeStyles, shorthands } from '@fluentui/react-components';
import { tokens } from '@fluentui/react-theme';

export const useProfileMenuStyles = makeStyles({
  trigger: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    ...shorthands.padding(0),
    ...shorthands.margin(0),
    ...shorthands.border('none'),
    backgroundColor: 'transparent',
    cursor: 'pointer',
    ...shorthands.borderRadius('50%'),
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
    ':active': {
      backgroundColor: tokens.colorNeutralBackground1Pressed,
    },
  },
  userInfo: {
    ...shorthands.padding('12px', '12px', '8px', '12px'),
    minWidth: '200px',
    maxWidth: '280px',
    overflow: 'hidden',
  },
  displayName: {
    display: 'block',
    ...shorthands.margin(0),
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  upn: {
    display: 'block',
    ...shorthands.margin('2px', 0, 0, 0),
    color: tokens.colorNeutralForeground2,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
});
