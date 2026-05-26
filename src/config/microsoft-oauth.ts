export interface MicrosoftOAuthConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  redirectUri: string;
}

export const MICROSOFT_IMAP_SCOPE = 'https://outlook.office.com/IMAP.AccessAsUser.All';
export const MICROSOFT_SMTP_SCOPE = 'https://outlook.office.com/SMTP.Send';
export const MICROSOFT_OAUTH_SCOPES = [
  MICROSOFT_IMAP_SCOPE,
  MICROSOFT_SMTP_SCOPE,
  'offline_access',
  'openid',
  'profile',
  'email',
].join(' ');

export function getMicrosoftOAuthConfig(port: number = 3000): MicrosoftOAuthConfig | null {
  const clientId = process.env.MICROSOFT_CLIENT_ID?.trim();
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET?.trim();

  if (!clientId || !clientSecret) {
    return null;
  }

  return {
    clientId,
    clientSecret,
    tenantId: process.env.MICROSOFT_TENANT_ID?.trim() || 'common',
    redirectUri:
      process.env.MICROSOFT_REDIRECT_URI?.trim() ||
      `http://localhost:${port}/api/oauth/microsoft/callback`,
  };
}

export function isMicrosoftOAuthConfigured(port: number = 3000): boolean {
  return getMicrosoftOAuthConfig(port) !== null;
}
