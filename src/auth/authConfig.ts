// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

/* global console, document*/

/// <reference types="office-js" />

import {
  BrowserAuthError,
  InteractionRequiredAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { msalConfig } from "./msalconfig";
import { createLocalUrl } from "./util";
import { getTokenRequest } from "./msalcommon";

export type AuthDialogResult = {
  accessToken?: string;
  error?: string;
};

type DialogEventMessage = { message: string; origin: string | undefined };
type DialogEventError = { error: number };
type DialogEventArg = DialogEventMessage | DialogEventError;

// Constants
const DIALOG_DIMENSIONS = { height: 60, width: 30 } as const;
const DIALOG_CLOSED_ERROR_CODE = 12006;
const POPUP_WINDOW_ERROR_CODE = "popup_window_error";
const SIGN_OUT_BUTTON_ID = "signOutButton";
const NESTED_APP_AUTH_REQUIREMENT = { name: "NestedAppAuth", version: "1.1" } as const;

// Encapsulate functions for getting user account and token information.
export class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private _dialogApiResult: Promise<string> | null = null;
  private _usingFallbackDialog = false;
  private readonly _boundSignOut = () => this.signOut();

  private setSignOutButtonVisibility(isVisible: boolean): void {
    const signOutButton = document.getElementById(SIGN_OUT_BUTTON_ID);
    if (signOutButton) {
      signOutButton.style.visibility = isVisible ? "visible" : "hidden";
    }
  }

  private isNestedAppAuthSupported(): boolean {
    return Office.context.requirements.isSetSupported(
      NESTED_APP_AUTH_REQUIREMENT.name, 
      NESTED_APP_AUTH_REQUIREMENT.version
    );
  }

  // Initialize MSAL public client application.
  async initialize(): Promise<void> {
    try {
      // Make sure office.js is initialized.
      await Office.onReady();

      // Initialize a nested public client application.
      this.pca = await createNestablePublicClientApplication(msalConfig);

      // If Office does not support nested app auth provide a sign-out button since the user selects account.
      if (!this.isNestedAppAuthSupported() && this.pca.getActiveAccount()) {
        this.setSignOutButtonVisibility(true);
      }
      
      // Add event listener for click event on sign out button.
      const signOutButton = document.getElementById(SIGN_OUT_BUTTON_ID);
      if (signOutButton) {
        signOutButton.addEventListener("click", this._boundSignOut);
      }
    } catch (error) {
      console.error("Failed to initialize AccountManager:", error);
      throw new Error(`Initialization failed: ${error}`);
    }
  }

  // Sign out the user.
  private async signOut() {
    await (this._usingFallbackDialog ? this.signOutWithDialogApi() : this.pca?.logoutPopup());
    this.setSignOutButtonVisibility(false);
  }

  // Get login hint for Word, Excel, or PowerPoint on the web from the auth context.
  private async getLoginHint(): Promise<string | undefined> {
    try {
      if (typeof Office !== "undefined" && Office.context) {
            const authContext = await Office.auth.getAuthContext();
            if (authContext?.userPrincipalName) return authContext.userPrincipalName;
        }
    } catch (error) {
      console.warn("Could not get login hint:", error);
    }
    return undefined;
  }

  async acquireToken(scopes: string[]): Promise<string> {
    // Check if the user is already signed in via fallback dialog API.
    if (this._dialogApiResult) {
      return this._dialogApiResult;
    }
    
    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }
    const loginHint = await this.getLoginHint();
    console.log(loginHint);
    
    try {
      console.log("Trying to acquire token silently...");
      const tokenRequest = getTokenRequest(scopes, false, undefined, loginHint);
      // If we have a login hint, use SSO silent flow which is required for Word, Excel, or PowerPoint on the web.
      const authResult = loginHint 
        ? await this.pca!.ssoSilent(tokenRequest)
        : await this.pca!.acquireTokenSilent(tokenRequest);
      console.log("Acquired token silently.");
      return authResult.accessToken;
    } catch (silentError) {
      if (silentError instanceof InteractionRequiredAuthError) {
        return this.acquireTokenInteractively(scopes, loginHint);
      } else if (silentError instanceof BrowserAuthError) {
        // NAA broker timed out or other browser auth issue (e.g. timed_out, popup_window_error).
        // This is expected in Excel Web where NAA may not respond reliably.
        console.warn(`Silent auth failed (${silentError.errorCode}), falling back to interactive`);
        return this.acquireTokenInteractively(scopes, loginHint);
      } else {
        throw new Error(`Unable to acquire access token: ${silentError}`);
      }
    }
  }

  private async acquireTokenInteractively(scopes: string[], loginHint: string | undefined): Promise<string> {
    try {
      console.log("Trying to acquire token interactively...");
      
      const authResult = await this.pca!.acquireTokenPopup(
        getTokenRequest(scopes, false, undefined, loginHint)
      );
      console.log("Acquired token interactively.");
      
      // Show sign-out button if Office doesn't support Nested App Auth
      if (!this.isNestedAppAuthSupported()) {
        this.setSignOutButtonVisibility(true);
      }
      return authResult.accessToken;
    } catch (popupError) {
      return this.handleInteractiveTokenError(popupError);
    }
  }

  private async handleInteractiveTokenError(popupError: unknown): Promise<string> {
    // Optional fallback if about:blank popup should not be shown
    if (popupError instanceof BrowserAuthError && popupError.errorCode === POPUP_WINDOW_ERROR_CODE) {
      const accessToken = await this.getTokenWithDialogApi();
      this.setSignOutButtonVisibility(true);
      return accessToken;
    } else {
      // Acquire token interactive failure.
      console.error(`Unable to acquire token interactively: ${popupError}`);
      throw new Error(`Unable to acquire access token: ${popupError}`);
    }
  }

  /**
   * Gets an access token by using the Office dialog API to handle authentication. Used for fallback scenario.
   * @returns The access token.
   */
  async getTokenWithDialogApi(): Promise<string> {
    this._dialogApiResult = new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html`), 
        { height: 60, width: 30 }, 
        (result: Office.AsyncResult<Office.Dialog>) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            this._dialogApiResult = null;
            reject(new Error(`displayDialogAsync failed: ${result.error?.message ?? 'unknown error'}`));
            return;
          }
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: DialogEventArg) => {
            if ((arg as DialogEventError).error === DIALOG_CLOSED_ERROR_CODE) {
              this._dialogApiResult = null;
              reject("Dialog closed");
            }
          });
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: DialogEventArg) => {
            let parsedMessage: { accessToken?: string; error?: string };
            try {
              parsedMessage = JSON.parse((arg as DialogEventMessage).message);
            } catch {
              dialog.close();
              this._dialogApiResult = null;
              reject(new Error("Failed to parse dialog message"));
              return;
            }
            dialog.close();
            if (parsedMessage.error) {
              this._dialogApiResult = null;
              reject(parsedMessage.error);
            } else if (parsedMessage.accessToken) {
              this.setSignOutButtonVisibility(true);
              this._usingFallbackDialog = true;
              resolve(parsedMessage.accessToken);
            } else {
              this._dialogApiResult = null;
              reject(new Error("Dialog message contained no access token"));
            }
          });
        }
      );
    });
    return this._dialogApiResult;
  }

  signOutWithDialogApi(): Promise<void> {
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        createLocalUrl(`dialog.html?logout=1`), 
        { height: 60, width: 30 }, 
        (result: Office.AsyncResult<Office.Dialog>) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.warn("signOutWithDialogApi: displayDialogAsync failed:", result.error?.message);
            reject(new Error(`displayDialogAsync failed: ${result.error?.message ?? 'unknown error'}`));
            return;
          }
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
            this.setSignOutButtonVisibility(false);
            this._dialogApiResult = null;
            dialog.close();
            resolve();
          });
        }
      );
    });
  }

  /**
   * Clean up resources and event listeners
   */
  cleanup(): void {
    const signOutButton = document.getElementById(SIGN_OUT_BUTTON_ID);
    if (signOutButton) {
      signOutButton.removeEventListener("click", this._boundSignOut);
    }
    this._dialogApiResult = null;
    this._usingFallbackDialog = false;
  }
}
