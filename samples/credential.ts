import type {
  AuthenticationProvider,
  RequestInformation,
} from "@microsoft/kiota-abstractions";
import {
  authorizationCodeGrant,
  buildAuthorizationUrl,
  calculatePKCECodeChallenge,
  discovery,
  randomPKCECodeVerifier,
  randomState,
  refreshTokenGrant,
  type Configuration,
  type TokenEndpointResponseHelpers,
} from "openid-client";
import type { TokenEndpointResponse } from "openid-client";
import * as http from "node:http";
import { spawn } from "node:child_process";

const OPEN_CMD =
  process.platform === "win32"
    ? "start"
    : process.platform === "darwin"
      ? "open"
      : "xdg-open";

const REDIRECT_PORT = 3000;
const REDIRECT_URI = `http://localhost:${REDIRECT_PORT}/callback`;

type TokenResponse = TokenEndpointResponse & TokenEndpointResponseHelpers;

export class BrowserInteractiveCredential implements AuthenticationProvider {
  private oidcConfig: Configuration | undefined;
  private cachedResponse: TokenResponse | undefined;

  constructor(
    private readonly params: {
      tenantId: string;
      clientId: string;
      scopes: string[];
    },
  ) {}

  async authenticateRequest(
    request: RequestInformation,
    _additionalAuthenticationContext?: Record<string, unknown>,
  ): Promise<void> {
    const token = await this.getAccessToken();
    request.headers.tryAdd("Authorization", `Bearer ${token}`);
  }

  private async getOidcConfig(): Promise<Configuration> {
    this.oidcConfig ??= await discovery(
      new URL(`https://login.microsoftonline.com/${this.params.tenantId}/v2.0`),
      this.params.clientId,
    );
    return this.oidcConfig;
  }

  private async getAccessToken(): Promise<string> {
    if (this.cachedResponse) {
      const remaining = this.cachedResponse.expiresIn();
      if (remaining === undefined || remaining > 60) {
        return this.cachedResponse.access_token;
      }
      if (this.cachedResponse.refresh_token) {
        const config = await this.getOidcConfig();
        this.cachedResponse = await refreshTokenGrant(
          config,
          this.cachedResponse.refresh_token,
          { scope: this.params.scopes.join(" ") },
        );
        return this.cachedResponse.access_token;
      }
    }

    const config = await this.getOidcConfig();
    this.cachedResponse = await this.acquireTokenInteractive(config);
    return this.cachedResponse.access_token;
  }

  private async acquireTokenInteractive(
    config: Configuration,
  ): Promise<TokenResponse> {
    const codeVerifier = randomPKCECodeVerifier();
    const codeChallenge = await calculatePKCECodeChallenge(codeVerifier);
    const state = randomState();

    const authUrl = buildAuthorizationUrl(config, {
      redirect_uri: REDIRECT_URI,
      scope: ["offline_access", ...this.params.scopes].join(" "),
      code_challenge: codeChallenge,
      code_challenge_method: "S256",
      state,
    });

    return new Promise((resolve, reject) => {
      const server = http.createServer(async (req, res) => {
        const callbackUrl = new URL(
          req.url!,
          `http://localhost:${REDIRECT_PORT}`,
        );
        if (!callbackUrl.searchParams.has("code")) {
          res.writeHead(204);
          res.end();
          return;
        }
        try {
          res.writeHead(200, { "Content-Type": "text/html" });
          res.end(
            "<h1>Login successful!</h1><p>You can close this window now.</p>",
          );
          server.close();
          const tokens = await authorizationCodeGrant(
            config,
            callbackUrl,
            { pkceCodeVerifier: codeVerifier, expectedState: state },
            { redirect_uri: REDIRECT_URI },
          );
          resolve(tokens);
        } catch (err) {
          reject(err);
        }
      });

      server.on("error", reject);
      server.listen(REDIRECT_PORT, () => {
        spawn(OPEN_CMD, [authUrl.href], { detached: true });
      });
    });
  }
}
