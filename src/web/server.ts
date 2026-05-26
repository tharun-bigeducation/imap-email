import express from 'express';
import cors from 'cors';
import bodyParser from 'body-parser';
import path from 'path';
import { fileURLToPath } from 'url';
import open from 'open';
import dotenv from 'dotenv';
import { AccountManager } from '../services/account-manager.js';
import { ImapService } from '../services/imap-service.js';
import { MicrosoftOAuthService } from '../services/microsoft-oauth-service.js';
import { TokenService } from '../services/token-service.js';
import { emailProviders, getProviderByEmail } from '../providers/email-providers.js';
import { ImapAccount } from '../types/index.js';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function sanitizeAccount(account: ImapAccount) {
  const { password, oauth, ...safeAccount } = account;
  return {
    ...safeAccount,
    authType: account.authType || 'password',
    oauthConfigured: account.authType === 'oauth2' && !!oauth,
    oauthExpiresAt: oauth?.expiresAt,
  };
}

export class WebUIServer {
  private app: express.Application;
  private accountManager: AccountManager;
  private imapService: ImapService;
  private microsoftOAuth: MicrosoftOAuthService;
  private tokenService: TokenService;
  private port: number;

  constructor(port: number = 3000) {
    this.app = express();
    this.port = port;
    this.accountManager = new AccountManager();
    this.imapService = new ImapService();
    this.microsoftOAuth = new MicrosoftOAuthService(port);
    this.tokenService = new TokenService(this.accountManager, this.microsoftOAuth);

    this.imapService.setAccountManager(this.accountManager);
    this.imapService.setTokenService(this.tokenService);
    
    this.setupMiddleware();
    this.setupRoutes();
  }

  private setupMiddleware(): void {
    this.app.use(cors());
    this.app.use(bodyParser.json());
    this.app.use(express.static(path.join(__dirname, '../../public')));
  }

  private setupRoutes(): void {
    // Get all providers
    this.app.get('/api/providers', (req, res) => {
      res.json(emailProviders);
    });

    this.app.get('/api/oauth/microsoft/config', (req, res) => {
      res.json(this.microsoftOAuth.getPublicConfig());
    });

    this.app.get('/api/oauth/microsoft/authorize', (req, res) => {
      try {
        const {
          accountName,
          email,
          providerId,
          host,
          port,
          tls,
          saveToSent,
        } = req.query;

        if (!accountName || !email || !providerId || !host) {
          res.status(400).json({ success: false, error: 'Missing required OAuth parameters' });
          return;
        }

        const authUrl = this.microsoftOAuth.createAuthorizationUrl({
          accountName: String(accountName),
          email: String(email),
          providerId: String(providerId),
          host: String(host),
          port: Number(port) || 993,
          tls: tls !== 'false',
          saveToSent: saveToSent !== 'false',
        });

        res.redirect(authUrl);
      } catch (error) {
        res.status(400).json({
          success: false,
          error: error instanceof Error ? error.message : 'Failed to start Microsoft OAuth flow',
        });
      }
    });

    this.app.get('/api/oauth/microsoft/callback', async (req, res) => {
      const { code, state, error, error_description: errorDescription } = req.query;

      if (error) {
        res.redirect(`/?oauth=error&message=${encodeURIComponent(String(errorDescription || error))}`);
        return;
      }

      if (!code || !state) {
        res.redirect('/?oauth=error&message=Missing%20OAuth%20callback%20parameters');
        return;
      }

      try {
        const { tokens, session } = await this.microsoftOAuth.exchangeAuthorizationCode(
          String(state),
          String(code)
        );

        const email = tokens.email || session.email;
        let account: ImapAccount;

        if (session.existingAccountId) {
          account = await this.accountManager.updateAccount(session.existingAccountId, {
            user: email,
            authType: 'oauth2',
            oauth: {
              provider: 'microsoft',
              accessToken: tokens.accessToken,
              refreshToken: tokens.refreshToken,
              expiresAt: tokens.expiresAt,
              scope: tokens.scope,
            },
          });
        } else {
          account = await this.accountManager.addAccount({
            name: session.accountName,
            host: session.host,
            port: session.port,
            user: email,
            password: '',
            tls: session.tls,
            authType: 'oauth2',
            oauth: {
              provider: 'microsoft',
              accessToken: tokens.accessToken,
              refreshToken: tokens.refreshToken,
              expiresAt: tokens.expiresAt,
              scope: tokens.scope,
            },
            saveToSent: session.saveToSent,
          });
        }

        const testResult = await this.imapService.testConnection(account);
        if (!testResult.success) {
          await this.accountManager.removeAccount(account.id);
          res.redirect(`/?oauth=error&message=${encodeURIComponent(testResult.error || 'Connection test failed')}`);
          return;
        }

        res.redirect(`/?oauth=success&accountId=${encodeURIComponent(account.id)}`);
      } catch (err) {
        res.redirect(`/?oauth=error&message=${encodeURIComponent(
          err instanceof Error ? err.message : 'OAuth sign-in failed'
        )}`);
      }
    });

    this.app.get('/api/oauth/microsoft/reauth/:id', (req, res) => {
      try {
        const account = this.accountManager.getAccount(req.params.id);
        if (!account || account.authType !== 'oauth2') {
          res.status(404).json({ success: false, error: 'OAuth account not found' });
          return;
        }

        const authUrl = this.microsoftOAuth.createAuthorizationUrl({
          accountName: account.name,
          email: account.user,
          providerId: 'office365',
          host: account.host,
          port: account.port,
          tls: account.tls,
          saveToSent: account.saveToSent !== false,
          existingAccountId: account.id,
        });

        res.redirect(authUrl);
      } catch (error) {
        res.status(400).json({
          success: false,
          error: error instanceof Error ? error.message : 'Failed to start re-authentication',
        });
      }
    });

    // Get all accounts
    this.app.get('/api/accounts', (req, res) => {
      try {
        const accounts = this.accountManager.getAllAccounts().map(sanitizeAccount);
        res.json(accounts);
      } catch (error) {
        res.status(500).json({ error: 'Failed to fetch accounts' });
      }
    });

    // Add new account
    this.app.post('/api/accounts', async (req, res) => {
      try {
        const { name, email, password, host, port, tls, smtp, imapUsername, authType } = req.body;
        
        if (authType !== 'oauth2' && !password) {
          res.status(400).json({ success: false, error: 'Password is required for non-OAuth accounts' });
          return;
        }

        // Auto-detect provider if not specified
        let imapHost = host;
        let imapPort = port;
        let useTls = tls;
        
        if (!host && email) {
          const provider = getProviderByEmail(email);
          if (provider) {
            imapHost = provider.imapHost;
            imapPort = provider.imapPort;
            useTls = provider.imapSecurity !== 'STARTTLS';
          }
        }
        
        const account = await this.accountManager.addAccount({
          name: name || email,
          host: imapHost,
          port: imapPort || 993,
          user: imapUsername || email,
          password: password || '',
          tls: useTls !== false,
          authType: authType || 'password',
          ...(imapUsername ? { email } : {}),
          smtp: smtp || undefined,
        });
        
        res.json({ success: true, account: sanitizeAccount(account) });
      } catch (error) {
        res.status(400).json({ 
          success: false, 
          error: error instanceof Error ? error.message : 'Failed to add account' 
        });
      }
    });

    // Test connection
    this.app.post('/api/test-connection', async (req, res) => {
      try {
        const { email, password, host, port, tls, imapUsername, authType, accessToken } = req.body;

        // Create temporary account for testing
        const testAccount: ImapAccount = {
          id: 'test-' + Date.now(),
          name: 'Test',
          host: host || 'imap.gmail.com',
          port: port || 993,
          user: imapUsername || email,
          password: password || '',
          tls: tls !== false,
          authType: authType || 'password',
          ...(authType === 'oauth2' && accessToken
            ? {
                oauth: {
                  provider: 'microsoft',
                  accessToken,
                  refreshToken: '',
                  expiresAt: Date.now() + 3600 * 1000,
                  scope: '',
                },
              }
            : {}),
        };
        
        const result = await this.imapService.testConnection(testAccount);
        
        if (!result.success) {
          res.status(400).json({
            success: false,
            error: result.error || 'Connection test failed',
          });
          return;
        }

        res.json({ 
          success: true, 
          folders: result.folders,
          messageCount: result.messageCount,
        });
      } catch (error) {
        res.status(400).json({ 
          success: false, 
          error: error instanceof Error ? error.message : 'Connection test failed' 
        });
      }
    });

    // Remove account
    this.app.delete('/api/accounts/:id', async (req, res) => {
      try {
        await this.accountManager.removeAccount(req.params.id);
        res.json({ success: true });
      } catch (error) {
        res.status(400).json({ 
          success: false, 
          error: error instanceof Error ? error.message : 'Failed to remove account' 
        });
      }
    });

    // Update account
    this.app.put('/api/accounts/:id', async (req, res) => {
      try {
        const { name, email, password, host, port, tls, smtp, saveToSent, imapUsername } = req.body;

        const updates: any = {};
        if (name !== undefined) updates.name = name;
        if (imapUsername) {
          updates.user = imapUsername;
          if (email !== undefined) updates.email = email;
        } else if (email !== undefined) {
          updates.user = email;
          updates.email = undefined;
        }
        if (password !== undefined) updates.password = password;
        if (host !== undefined) updates.host = host;
        if (port !== undefined) updates.port = port;
        if (tls !== undefined) updates.tls = tls;
        if (smtp !== undefined) updates.smtp = smtp;
        if (saveToSent !== undefined) updates.saveToSent = saveToSent;
        
        const account = await this.accountManager.updateAccount(req.params.id, updates);
        res.json({ success: true, account: sanitizeAccount(account) });
      } catch (error) {
        res.status(400).json({ 
          success: false, 
          error: error instanceof Error ? error.message : 'Failed to update account' 
        });
      }
    });

    // Get single account
    this.app.get('/api/accounts/:id', async (req, res) => {
      try {
        const account = this.accountManager.getAccount(req.params.id);
        if (!account) {
          res.status(404).json({ success: false, error: 'Account not found' });
        } else {
          res.json({ success: true, account: sanitizeAccount(account) });
        }
      } catch (error) {
        res.status(400).json({ 
          success: false, 
          error: error instanceof Error ? error.message : 'Failed to get account' 
        });
      }
    });

    // Test existing account connection (without re-entering password)
    this.app.post('/api/accounts/:id/test', async (req, res) => {
      try {
        const account = this.accountManager.getAccount(req.params.id);
        if (!account) {
          res.status(404).json({ success: false, error: 'Account not found' });
          return;
        }

        const result = await this.imapService.testConnection(account);

        res.json({
          success: result.success,
          accountId: account.id,
          accountName: account.name,
          host: account.host,
          folders: result.folders,
          messageCount: result.messageCount,
          error: result.error,
        });
      } catch (error) {
        res.status(400).json({
          success: false,
          error: error instanceof Error ? error.message : 'Test failed',
        });
      }
    });

    // Health check
    this.app.get('/api/health', (req, res) => {
      res.json({
        status: 'ok',
        version: '1.0.0',
        microsoftOAuthConfigured: this.microsoftOAuth.isConfigured(),
      });
    });
  }

  async start(autoOpen: boolean = true): Promise<void> {
    return new Promise((resolve) => {
      const server = this.app.listen(this.port, () => {
        console.log(`🌐 Web UI server running at http://localhost:${this.port}`);
        if (this.microsoftOAuth.isConfigured()) {
          console.log('✓ Microsoft OAuth configured');
        } else {
          console.log('ℹ Microsoft OAuth not configured — add MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET to .env for company accounts');
        }
        
        if (autoOpen) {
          // Open browser after a short delay
          setTimeout(() => {
            open(`http://localhost:${this.port}`);
          }, 1000);
        }
        
        resolve();
      });

      // Handle graceful shutdown
      process.on('SIGINT', () => {
        console.log('\nShutting down web server...');
        server.close(() => {
          process.exit(0);
        });
      });
    });
  }
}

// CLI entry point
if (import.meta.url === `file://${process.argv[1]}`) {
  const port = parseInt(process.env.PORT || '3000');
  const server = new WebUIServer(port);
  server.start();
}
