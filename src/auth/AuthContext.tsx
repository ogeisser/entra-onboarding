import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from 'react';
import type { AccountManager } from '@/auth/authConfig';

export interface AuthUser {
  displayName: string;
  userPrincipalName: string;
}

export interface AuthContextValue {
  /** Whether the host supports Nested App Authentication (Office requirement set) */
  isNaaSupported: boolean;
  /** Whether NAA is available and initialized; login is only allowed when true */
  naaAvailable: boolean;
  /** Last NAA init error, if any (e.g. when naaAvailable is false due to init failure) */
  naaInitError: Error | null;
  /** Currently logged-in user, or null if not logged in */
  user: AuthUser | null;
  /** Whether login() is in progress */
  loginInProgress: boolean;
  /** Perform SSO login (get token + fetch /me, set user on success) */
  login: () => Promise<void>;
  /** Perform interactive login with account selection (popup, prompt: select_account) */
  loginAdmin: () => Promise<void>;
  /** Clear user and error (return to login view) */
  logout: () => void;
  /** Get an access token for the given scopes. Defaults to ['User.ReadWrite.All'] */
  getAccessToken: (scopes?: string[]) => Promise<string>;
  /** Last auth error, if any */
  error: Error | null;
  /** Clear the last error */
  clearError: () => void;
}

const AuthContext = createContext<AuthContextValue | null>(null);

export function useAuth(): AuthContextValue {
  const ctx = useContext(AuthContext);
  if (!ctx) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return ctx;
}

function isNaaSupported(): boolean {
  try {
    if (typeof Office !== 'undefined' && Office.context?.requirements?.isSetSupported) {
      return Office.context.requirements.isSetSupported('NestedAppAuth', '1.1');
    }
  } catch {
    // ignore
  }
  return false;
}

interface AuthProviderProps {
  children: ReactNode;
  accountManager: AccountManager;
}

export function AuthProvider({ children, accountManager }: AuthProviderProps) {
  const [error, setError] = useState<Error | null>(null);
  const [user, setUser] = useState<AuthUser | null>(null);
  const [loginInProgress, setLoginInProgress] = useState(false);
  const [naaAvailable, setNaaAvailable] = useState(false);
  const [naaInitError, setNaaInitError] = useState<Error | null>(null);
  const naaSupported = useMemo(() => isNaaSupported(), []);

  useEffect(() => {
    let cancelled = false;
    console.log('Auth: initializing AccountManager');
    accountManager.initialize().then(
      () => {
        if (cancelled) return;
        setNaaAvailable(true);
        setNaaInitError(null);
        console.log('Auth: AccountManager ready');
      },
      (err) => {
        if (cancelled) return;
        const e = err instanceof Error ? err : new Error(String(err));
        setNaaAvailable(false);
        setNaaInitError(e);
        console.error('Auth: AccountManager init failed', e);
      }
    );
    return () => {
      cancelled = true;
    };
  }, [accountManager]);

  const getAccessToken = useCallback(
    async (scopes?: string[]) => {
      setError(null);
      try {
        return await accountManager.acquireToken(scopes ?? ['User.ReadWrite.All']);
      } catch (err) {
        const e = err instanceof Error ? err : new Error(String(err));
        setError(e);
        throw e;
      }
    },
    [accountManager]
  );

  const clearError = useCallback(() => setError(null), []);

  const login = useCallback(async () => {
    clearError();
    setLoginInProgress(true);
    console.log('Auth: login started');
    try {
      const token = await getAccessToken(['User.ReadWrite.All']);
      const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!response.ok) {
        const text = await response.text();
        const e = new Error(text || `Graph error ${response.status}`);
        setError(e);
        throw e;
      }
      const data = (await response.json()) as {
        displayName?: string;
        userPrincipalName?: string;
      };
      setUser({
        displayName: data.displayName ?? '',
        userPrincipalName: data.userPrincipalName ?? '',
      });
      console.log('Auth: login success');
    } catch (err) {
      console.error('Auth: login failed', err);
      setError(err instanceof Error ? err : new Error(String(err)));
      // Do not rethrow: avoids unhandled promise rejection when e.g. user cancels sign-in
    } finally {
      setLoginInProgress(false);
    }
  }, [clearError, getAccessToken]);

  const loginAdmin = useCallback(async () => {
    clearError();
    setLoginInProgress(true);
    console.log('Auth: admin login started');
    try {
      const token = await accountManager.acquireTokenWithAccountSelection(['User.ReadWrite.All']);
      const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!response.ok) {
        const text = await response.text();
        const e = new Error(text || `Graph error ${response.status}`);
        setError(e);
        throw e;
      }
      const data = (await response.json()) as {
        displayName?: string;
        userPrincipalName?: string;
      };
      setUser({
        displayName: data.displayName ?? '',
        userPrincipalName: data.userPrincipalName ?? '',
      });
      console.log('Auth: admin login success');
    } catch (err) {
      console.error('Auth: admin login failed', err);
      setError(err instanceof Error ? err : new Error(String(err)));
    } finally {
      setLoginInProgress(false);
    }
  }, [clearError, accountManager]);

  const logout = useCallback(() => {
    setUser(null);
    clearError();
  }, [clearError]);

  const value = useMemo<AuthContextValue>(
    () => ({
      isNaaSupported: naaSupported,
      naaAvailable,
      naaInitError,
      user,
      loginInProgress,
      login,
      loginAdmin,
      logout,
      getAccessToken,
      error,
      clearError,
    }),
    [naaSupported, naaAvailable, naaInitError, user, loginInProgress, login, loginAdmin, logout, getAccessToken, error, clearError]
  );

  return (
    <AuthContext.Provider value={value}>{children}</AuthContext.Provider>
  );
}
