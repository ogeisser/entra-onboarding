import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.tsx';
import { AccountManager } from "@/auth/authConfig";

console.log('Taskpane main.tsx loaded – test for Bun remote console relay');
import { AuthProvider } from '@/auth/AuthContext.tsx';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

/* global Office */

const elem = document.getElementById('root')!;
const accountManager = new AccountManager();

const renderApp = (root: ReactDOM.Root) => {
  root.render(
    <React.StrictMode>
      <FluentProvider theme={webLightTheme}>
        <AuthProvider accountManager={accountManager}>
          <App />
        </AuthProvider>
      </FluentProvider>
    </React.StrictMode>,
  );
};

Office.onReady(() => {
  if (import.meta.hot) {
    // With hot module reloading, `import.meta.hot.data` is persisted.
    const root = (import.meta.hot.data.root ??= ReactDOM.createRoot(elem));
    renderApp(root);
  } else {
    // The hot module reloading API is not available in production.
    const root = ReactDOM.createRoot(elem);
    renderApp(root);
  }
});
