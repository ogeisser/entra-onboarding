import {
  Avatar,
  Body1Strong,
  Caption1,
  Menu,
  MenuDivider,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
} from '@fluentui/react-components';
import { SignOut20Regular } from '@fluentui/react-icons';
import type { AuthUser } from '../auth/AuthContext';
import { useProfileMenuStyles } from './ProfileMenu.styles';

interface ProfileMenuProps {
  user: AuthUser;
  onLogout: () => void;
}

export function ProfileMenu({ user, onLogout }: ProfileMenuProps) {
  const classes = useProfileMenuStyles();

  return (
    <Menu>
      <MenuTrigger>
        <button
          type="button"
          aria-label="Open profile menu"
          className={classes.trigger}
        >
          <Avatar name={user.displayName} size={32} />
        </button>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <div className={classes.userInfo}>
            <Body1Strong className={classes.displayName}>{user.displayName}</Body1Strong>
            <Caption1 className={classes.upn}>{user.userPrincipalName}</Caption1>
          </div>
          <MenuDivider />
          <MenuItem icon={<SignOut20Regular />} onClick={onLogout}>
            Sign out
          </MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
