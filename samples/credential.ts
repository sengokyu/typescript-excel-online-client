import { AccessToken, GetTokenOptions, TokenCredential } from "@azure/identity";
import {
  AccountInfo,
  AuthenticationResult,
  LogLevel,
  PublicClientApplication,
} from "@azure/msal-node";
import * as child_process from "child_process";

const OPEN_CMD =
  process.platform === "win32"
    ? "start"
    : process.platform === "darwin"
      ? "open"
      : "xdg-open";

export class BrowserInteractiveCredential implements TokenCredential {
  private readonly clientApplication: PublicClientApplication;

  constructor({ tenantId, clientId }: { tenantId: string; clientId: string }) {
    this.clientApplication = new PublicClientApplication({
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      system: {
        loggerOptions: {
          loggerCallback(loglevel, message, containsPii) {
            console.log(message);
          },
          piiLoggingEnabled: false,
          logLevel: LogLevel.Error,
        },
      },
    });
  }

  async getToken(
    scopes: string | string[],
    options?: GetTokenOptions,
  ): Promise<AccessToken | null> {
    const accounts = await this.clientApplication
      .getTokenCache()
      .getAllAccounts();

    const scopeArray = Array.isArray(scopes) ? scopes : [scopes];
    const authResult =
      !accounts || accounts.length === 0
        ? await this.getTokenInteractive(scopeArray)
        : await this.getTokenSilent(scopeArray, accounts[0] as AccountInfo);

    return {
      token: authResult.accessToken,
      expiresOnTimestamp: authResult.expiresOn!.getTime(),
    };
  }

  private async getTokenInteractive(
    scopes: string[],
  ): Promise<AuthenticationResult> {
    return await this.clientApplication.acquireTokenInteractive({
      scopes: scopes,
      openBrowser: async (url) => {
        child_process.spawn(OPEN_CMD, [url], { detached: true });
      },
      successTemplate:
        "<h1>Login successful!</h1><p>You can close this window now.</p>",
      errorTemplate:
        "<h1>Login failed</h1><p>Something went wrong during login. Please try again.</p>",
    });
  }

  private async getTokenSilent(
    scopes: string[],
    account: AccountInfo,
  ): Promise<AuthenticationResult> {
    return await this.clientApplication.acquireTokenSilent({ scopes, account });
  }
}
