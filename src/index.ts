import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import { ImapService } from './services/imap-service.js';
import { AccountManager } from './services/account-manager.js';
import { SmtpService } from './services/smtp-service.js';
import { SpamService } from './services/spam-service.js';
import { MicrosoftOAuthService } from './services/microsoft-oauth-service.js';
import { TokenService } from './services/token-service.js';
import { registerTools } from './tools/index.js';

// Silence any package version output to stdout
const originalWrite = process.stdout.write.bind(process.stdout);
(process.stdout.write as any) = function(chunk: any, encoding?: any, callback?: any): boolean {
  // Only allow JSON-RPC messages through
  if (typeof chunk === 'string' && (chunk.startsWith('{') || chunk === '\n')) {
    return originalWrite(chunk, encoding, callback);
  }
  return true;
};

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
dotenv.config({ path: path.join(projectRoot, '.env') });

const server = new McpServer({
  name: 'imap-mcp-server',
  version: '1.0.0',
});

const imapService = new ImapService();
const accountManager = new AccountManager();
const smtpService = new SmtpService();
const spamService = new SpamService();
const microsoftOAuth = new MicrosoftOAuthService(parseInt(process.env.PORT || '3000'));
const tokenService = new TokenService(accountManager, microsoftOAuth);

// Allow ImapService to auto-connect using stored credentials
imapService.setAccountManager(accountManager);
imapService.setTokenService(tokenService);
smtpService.setTokenService(tokenService);

// Register all tools
registerTools(server, imapService, accountManager, smtpService, spamService);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('IMAP MCP Server started');
}

main().catch((error) => {
  console.error('Server error:', error);
  process.exit(1);
});