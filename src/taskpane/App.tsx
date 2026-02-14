import { useCallback, useEffect, useState, type ReactNode } from 'react';
import type { SelectTabData, TabValue } from '@fluentui/react-components';
import {
  Body1,
  Button,
  Card,
  CardFooter,
  CardHeader,
  Tab,
  TabList,
  Title3,
} from '@fluentui/react-components';
import { ProfileMenu } from './components/ProfileMenu';
import { CreateTab } from './components/CreateTab';
import { UpdateTab } from './components/UpdateTab';
import { useAuth } from '@/auth/AuthContext.tsx';
import { useAppStyles } from './App.styles';

/** Formatted error details for auth debugging (incl. NAA/Bridge errors). */
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
    isNaaSupported,
    naaAvailable,
    naaInitError,
    user,
    loginInProgress,
    login,
    loginAdmin,
    logout,
    error: authError,
    clearError,
  } = useAuth();
  const [selectedTab, setSelectedTab] = useState<TabValue>('create');
  const [loginHint, setLoginHint] = useState<string | null>(null);

  useEffect(() => {
    if (user !== null || !naaAvailable) return;
    const hasGetAuthContext = typeof (Office as unknown as { auth?: { getAuthContext?: unknown } })?.auth?.getAuthContext === 'function';
    if (!hasGetAuthContext) return;
    Office.auth.getAuthContext().then((ctx) => {
      const hint = ctx.userPrincipalName ?? ctx.loginHint ?? null;
      setLoginHint(hint || null);
    }).catch(() => setLoginHint(null));
  }, [user, naaAvailable]);

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
    } catch {
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

  // ─── Sign-in screen ────────────────────────────────────────────────

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
                    ? <>Sign in using your Office account.{loginHint ? <><br /><br /><strong>{loginHint}</strong></> : null}</>
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
          <Card>
            <CardHeader
              header={<Title3>Admin Sign In</Title3>}
              description={<Body1>Sign in using your Admin account.</Body1>}
            />
            <CardFooter
              action={
                <Button
                  appearance="primary"
                  onClick={loginAdmin}
                  disabled={!naaAvailable || loginInProgress}
                >
                  {loginInProgress ? 'Signing in…' : 'Sign In'}
                </Button>
              }
            />
          </Card>
        </div>
      </div>
    );
  }

  // ─── Main app (authenticated) ──────────────────────────────────────

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
        {selectedTab === 'create' && <CreateTab />}
        {selectedTab === 'update' && <UpdateTab />}
      </div>
    </div>
  );
}

export default App;
