// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file handles MSAL auth for the fallback dialog page. */

/* global Office, window, URLSearchParams */

import {
    type AuthenticationResult,
    createStandardPublicClientApplication,
    type IPublicClientApplication,
  } from "@azure/msal-browser";
  import { getTokenRequest } from "@/auth/msalcommon";
  import { createLocalUrl } from "@/auth/util";
  import { defaultScopes, msalConfig } from "@/auth/msalconfig";
  import type { AuthDialogResult } from "@/auth/authConfig";
  
  // read querystring parameter
  function getQueryParameter(param: string) {
    const params = new URLSearchParams(window.location.search);
    return params.get(param);
  }
  
  async function sendDialogMessage(message: string) {
    await Office.onReady();
    Office.context.ui.messageParent(message);
  }
  async function returnResult(publicClientApp: IPublicClientApplication, authResult: AuthenticationResult) {
    publicClientApp.setActiveAccount(authResult.account);
  
    const authDialogResult: AuthDialogResult = {
      accessToken: authResult.accessToken,
    };
  
    sendDialogMessage(JSON.stringify(authDialogResult));
  }
  
  export async function initializeMsal() {
    // Use standard Public Client instead of nested because this is a fallback path when nested app authentication isn't available.
    const publicClientApp = await createStandardPublicClientApplication(msalConfig);

    // Admin mode: skip silent acquisition and always show the account picker.
    const isAdminLogin = getQueryParameter("admin") === "1";

    try {
      if (getQueryParameter("logout") === "1") {
        await publicClientApp.logoutRedirect({ postLogoutRedirectUri: createLocalUrl("dialog.html?close=1") });
        return;
      } else if (getQueryParameter("close") === "1") {
        sendDialogMessage("close");
        return;
      }
      const result = await publicClientApp.handleRedirectPromise();
  
      if (result) {
        return returnResult(publicClientApp, result);
      }
    } catch (ex: any) {
      const authDialogResult: AuthDialogResult = {
        error: ex.name,
      };
      sendDialogMessage(JSON.stringify(authDialogResult));
      return;
    }
  
    // In admin mode we always go straight to the redirect with account picker (prompt: select_account).
    // In normal mode we first try silent token acquisition.
    if (!isAdminLogin) {
      try {
        if (publicClientApp.getActiveAccount()) {
          const result = await publicClientApp.acquireTokenSilent(getTokenRequest(defaultScopes, false));
          if (result) {
            return returnResult(publicClientApp, result);
          }
        }
      } catch {
        /* empty */
      }
    }
  
    // Redirect to Azure AD. In admin mode selectAccount is true so the account picker is shown.
    // The redirectUri points to dialog.html (without query params) so no extra Azure App Registration entry is needed.
    publicClientApp.acquireTokenRedirect(getTokenRequest(defaultScopes, true, createLocalUrl("dialog.html")));
  }
  
  initializeMsal();