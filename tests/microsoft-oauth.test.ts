import { describe, it, expect, vi, beforeEach } from 'vitest';
import { MicrosoftOAuthService } from '../src/services/microsoft-oauth-service.js';

describe('MicrosoftOAuthService', () => {
  beforeEach(() => {
    vi.unstubAllGlobals();
    delete process.env.MICROSOFT_CLIENT_ID;
    delete process.env.MICROSOFT_CLIENT_SECRET;
    delete process.env.MICROSOFT_TENANT_ID;
    delete process.env.MICROSOFT_REDIRECT_URI;
  });

  it('should report not configured without env vars', () => {
    const service = new MicrosoftOAuthService(3000);
    expect(service.isConfigured()).toBe(false);
    expect(service.getPublicConfig()).toEqual({ configured: false });
  });

  it('should build authorization URL when configured', () => {
    process.env.MICROSOFT_CLIENT_ID = 'test-client-id';
    process.env.MICROSOFT_CLIENT_SECRET = 'test-client-secret';

    const service = new MicrosoftOAuthService(3000);
    const url = service.createAuthorizationUrl({
      accountName: 'Work Email',
      email: 'user@company.com',
      providerId: 'office365',
      host: 'outlook.office365.com',
      port: 993,
      tls: true,
      saveToSent: true,
    });

    expect(url).toContain('login.microsoftonline.com');
    expect(url).toContain('client_id=test-client-id');
    expect(url).toContain('code_challenge=');
    expect(url).toContain('login_hint=user%40company.com');
  });

  it('should exchange authorization code for tokens', async () => {
    process.env.MICROSOFT_CLIENT_ID = 'test-client-id';
    process.env.MICROSOFT_CLIENT_SECRET = 'test-client-secret';

    const payload = Buffer.from(JSON.stringify({ preferred_username: 'user@company.com' })).toString('base64url');
    const idToken = `header.${payload}.signature`;

    vi.stubGlobal('fetch', vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({
        access_token: 'access-token',
        refresh_token: 'refresh-token',
        expires_in: 3600,
        scope: 'imap',
        id_token: idToken,
      }),
    }));

    const service = new MicrosoftOAuthService(3000);
    const authUrl = service.createAuthorizationUrl({
      accountName: 'Work Email',
      email: 'user@company.com',
      providerId: 'office365',
      host: 'outlook.office365.com',
      port: 993,
      tls: true,
      saveToSent: true,
    });

    const state = new URL(authUrl).searchParams.get('state');
    const result = await service.exchangeAuthorizationCode(state!, 'auth-code');

    expect(result.tokens.accessToken).toBe('access-token');
    expect(result.tokens.refreshToken).toBe('refresh-token');
    expect(result.tokens.email).toBe('user@company.com');
    expect(result.session.accountName).toBe('Work Email');
  });
});
