import crypto from 'crypto';
import {
  getMicrosoftOAuthConfig,
  MICROSOFT_OAUTH_SCOPES,
  type MicrosoftOAuthConfig,
} from '../config/microsoft-oauth.js';

export interface MicrosoftTokenResponse {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  scope: string;
  email?: string;
}

interface PendingOAuthSession {
  codeVerifier: string;
  accountName: string;
  email: string;
  providerId: string;
  host: string;
  port: number;
  tls: boolean;
  saveToSent: boolean;
  existingAccountId?: string;
  expiresAt: number;
}

function decodeJwtPayload(token: string): Record<string, unknown> {
  const payload = token.split('.')[1];
  if (!payload) {
    return {};
  }

  const normalized = payload.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), '=');
  return JSON.parse(Buffer.from(padded, 'base64').toString('utf8'));
}

export class MicrosoftOAuthService {
  private pendingSessions = new Map<string, PendingOAuthSession>();
  private readonly port: number;

  constructor(port: number = 3000) {
    this.port = port;
  }

  isConfigured(): boolean {
    return getMicrosoftOAuthConfig(this.port) !== null;
  }

  getPublicConfig(): { configured: boolean; redirectUri?: string; tenantId?: string } {
    const config = getMicrosoftOAuthConfig(this.port);
    if (!config) {
      return { configured: false };
    }

    return {
      configured: true,
      redirectUri: config.redirectUri,
      tenantId: config.tenantId,
    };
  }

  createAuthorizationUrl(input: {
    accountName: string;
    email: string;
    providerId: string;
    host: string;
    port: number;
    tls: boolean;
    saveToSent: boolean;
    existingAccountId?: string;
  }): string {
    const config = this.requireConfig();
    const state = crypto.randomBytes(16).toString('hex');
    const codeVerifier = crypto.randomBytes(32).toString('base64url');
    const codeChallenge = crypto
      .createHash('sha256')
      .update(codeVerifier)
      .digest('base64url');

    this.pendingSessions.set(state, {
      codeVerifier,
      accountName: input.accountName,
      email: input.email,
      providerId: input.providerId,
      host: input.host,
      port: input.port,
      tls: input.tls,
      saveToSent: input.saveToSent,
      existingAccountId: input.existingAccountId,
      expiresAt: Date.now() + 10 * 60 * 1000,
    });

    this.cleanupExpiredSessions();

    const params = new URLSearchParams({
      client_id: config.clientId,
      response_type: 'code',
      redirect_uri: config.redirectUri,
      response_mode: 'query',
      scope: MICROSOFT_OAUTH_SCOPES,
      state,
      code_challenge: codeChallenge,
      code_challenge_method: 'S256',
      login_hint: input.email,
    });

    return `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/authorize?${params.toString()}`;
  }

  async exchangeAuthorizationCode(state: string, code: string): Promise<{
    tokens: MicrosoftTokenResponse;
    session: PendingOAuthSession;
  }> {
    const session = this.consumePendingSession(state);
    const config = this.requireConfig();
    const tokens = await this.requestTokens(config, {
      grant_type: 'authorization_code',
      code,
      redirect_uri: config.redirectUri,
      code_verifier: session.codeVerifier,
    });

    return { tokens, session };
  }

  async refreshAccessToken(refreshToken: string): Promise<MicrosoftTokenResponse> {
    const config = this.requireConfig();
    return this.requestTokens(config, {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      scope: MICROSOFT_OAUTH_SCOPES,
    });
  }

  private consumePendingSession(state: string): PendingOAuthSession {
    const session = this.pendingSessions.get(state);
    this.pendingSessions.delete(state);

    if (!session) {
      throw new Error('OAuth session expired or invalid. Please try signing in again.');
    }

    if (session.expiresAt < Date.now()) {
      throw new Error('OAuth session expired. Please try signing in again.');
    }

    return session;
  }

  private cleanupExpiredSessions(): void {
    const now = Date.now();
    for (const [state, session] of this.pendingSessions.entries()) {
      if (session.expiresAt < now) {
        this.pendingSessions.delete(state);
      }
    }
  }

  private requireConfig(): MicrosoftOAuthConfig {
    const config = getMicrosoftOAuthConfig(this.port);
    if (!config) {
      throw new Error(
        'Microsoft OAuth is not configured. Set MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET in a .env file.'
      );
    }
    return config;
  }

  private async requestTokens(
    config: MicrosoftOAuthConfig,
    body: Record<string, string>
  ): Promise<MicrosoftTokenResponse> {
    const params = new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      ...body,
    });

    const response = await fetch(
      `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString(),
      }
    );

    const data = await response.json() as {
      access_token?: string;
      refresh_token?: string;
      expires_in?: number;
      scope?: string;
      id_token?: string;
      error?: string;
      error_description?: string;
    };

    if (!response.ok || !data.access_token) {
      throw new Error(data.error_description || data.error || 'Failed to obtain Microsoft OAuth tokens');
    }

    let email: string | undefined;
    if (data.id_token) {
      const claims = decodeJwtPayload(data.id_token);
      email =
        (claims.preferred_username as string | undefined) ||
        (claims.email as string | undefined) ||
        (claims.upn as string | undefined);
    }

    return {
      accessToken: data.access_token,
      refreshToken: data.refresh_token || body.refresh_token || '',
      expiresAt: Date.now() + (data.expires_in || 3600) * 1000,
      scope: data.scope || MICROSOFT_OAUTH_SCOPES,
      email,
    };
  }
}
