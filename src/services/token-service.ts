import type { AccountManager } from './account-manager.js';
import type { MicrosoftOAuthService } from './microsoft-oauth-service.js';
import type { ImapAccount } from '../types/index.js';

export class TokenService {
  private refreshPromises = new Map<string, Promise<string>>();

  constructor(
    private accountManager: AccountManager,
    private microsoftOAuth: MicrosoftOAuthService
  ) {}

  async getAccessToken(account: ImapAccount): Promise<string> {
    if (account.authType !== 'oauth2' || !account.oauth) {
      return account.password;
    }

    const bufferMs = 5 * 60 * 1000;
    if (account.oauth.expiresAt > Date.now() + bufferMs) {
      return account.oauth.accessToken;
    }

    const existingRefresh = this.refreshPromises.get(account.id);
    if (existingRefresh) {
      return existingRefresh;
    }

    const refreshPromise = this.refreshAccessToken(account);
    this.refreshPromises.set(account.id, refreshPromise);

    try {
      return await refreshPromise;
    } finally {
      this.refreshPromises.delete(account.id);
    }
  }

  private async refreshAccessToken(account: ImapAccount): Promise<string> {
    if (!account.oauth?.refreshToken) {
      throw new Error(`OAuth refresh token missing for account ${account.name}. Please sign in again.`);
    }

    if (account.oauth.provider !== 'microsoft') {
      throw new Error(`Unsupported OAuth provider: ${account.oauth.provider}`);
    }

    const tokens = await this.microsoftOAuth.refreshAccessToken(account.oauth.refreshToken);
    await this.accountManager.updateOAuthTokens(account.id, {
      accessToken: tokens.accessToken,
      refreshToken: tokens.refreshToken || account.oauth.refreshToken,
      expiresAt: tokens.expiresAt,
      scope: tokens.scope,
    });

    return tokens.accessToken;
  }
}
