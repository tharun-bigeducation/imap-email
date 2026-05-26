// Email providers data
let providers = [];
let selectedProvider = null;
let currentStep = 1;
let microsoftOAuthConfig = { configured: false };
let usePasswordAuth = false;

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
    await loadProviders();
    await loadMicrosoftOAuthConfig();
    renderProviders();
    handleOAuthRedirect();
    
    // Setup form handler
    document.getElementById('accountForm').onsubmit = handleAccountSubmit;
    
    // Clear test results when form fields change
    const formFields = ['email', 'password', 'imapHost', 'imapPort'];
    formFields.forEach(fieldId => {
        document.getElementById(fieldId).addEventListener('input', () => {
            document.getElementById('inlineTestResult').classList.add('hidden');
            document.getElementById('inlineTestSuccess').classList.add('hidden');
            document.getElementById('inlineTestError').classList.add('hidden');
        });
    });
});

async function loadMicrosoftOAuthConfig() {
    try {
        const response = await fetch('/api/oauth/microsoft/config');
        microsoftOAuthConfig = await response.json();
    } catch (error) {
        console.error('Failed to load Microsoft OAuth config:', error);
    }
}

function handleOAuthRedirect() {
    const params = new URLSearchParams(window.location.search);
    const oauthStatus = params.get('oauth');

    if (!oauthStatus) {
        return;
    }

    window.history.replaceState({}, document.title, window.location.pathname);

    if (oauthStatus === 'success') {
        if (params.get('return') === 'accounts') {
            viewAccounts('Microsoft account re-authenticated successfully. Restart Claude Desktop to use the new tokens.');
        } else {
            goToStep(3);
            document.getElementById('testProgress').classList.add('hidden');
            document.getElementById('testError').classList.add('hidden');
            document.getElementById('testSuccess').classList.remove('hidden');
        }
        return;
    }

    if (oauthStatus === 'error') {
        const message = params.get('message') || 'Microsoft sign-in failed';
        if (params.get('return') === 'accounts') {
            viewAccounts('Re-authentication failed: ' + message, 'error');
        } else {
            goToStep(2);
            document.getElementById('inlineTestResult').classList.remove('hidden');
            document.getElementById('inlineTestSuccess').classList.add('hidden');
            document.getElementById('inlineTestError').classList.remove('hidden');
            document.getElementById('inlineErrorMessage').textContent = '❌ ' + message;
        }
    }
}

function showAccountsAlert(message, type = 'success') {
    const alert = document.getElementById('accountsListAlert');
    if (!alert) return;

    const styles = type === 'error'
        ? 'bg-red-50 border-red-200 text-red-800'
        : 'bg-green-50 border-green-200 text-green-800';

    alert.className = `mb-4 border rounded-md p-4 text-sm ${styles}`;
    alert.textContent = message;
    alert.classList.remove('hidden');
}

function isMicrosoftOAuthAccount(account) {
    return account.authType === 'oauth2' ||
        /outlook\.office365\.com|imap-mail\.outlook\.com/i.test(account.host || '');
}

// Load providers from API
async function loadProviders() {
    try {
        const response = await fetch('/api/providers');
        providers = await response.json();
    } catch (error) {
        console.error('Failed to load providers:', error);
    }
}

// Render provider grid
function renderProviders() {
    const grid = document.getElementById('providerGrid');
    grid.innerHTML = providers.map(provider => `
        <div class="provider-card bg-white rounded-lg shadow p-4 cursor-pointer hover:shadow-lg border-l-4" onclick="selectProvider('${provider.id}')" style="border-left-color: ${provider.color}">
            <div class="text-center">
                <div class="h-12 w-12 mx-auto mb-2 rounded-lg flex items-center justify-center p-2" style="background-color: ${provider.color}15;">
                    <img src="${provider.iconUrl}" alt="${provider.name}" class="w-full h-full object-contain" style="filter: brightness(0) saturate(100%) invert(0)" 
                         onload="this.style.filter = 'none'" 
                         onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';">
                    <div class="w-full h-full items-center justify-center text-sm font-bold text-white rounded" style="display:none; background-color: ${provider.color};">
                        ${provider.name.charAt(0)}
                    </div>
                </div>
                <h3 class="font-medium text-gray-900">${provider.displayName}</h3>
                ${provider.domains.length > 0 ? `<p class="text-sm text-gray-500 mt-1">${provider.domains[0]}</p>` : ''}
            </div>
        </div>
    `).join('');
}

// Select provider
function selectProvider(providerId) {
    selectedProvider = providers.find(p => p.id === providerId);
    usePasswordAuth = selectedProvider?.id !== 'office365';
    goToStep(2);
    
    // Pre-fill advanced settings
    if (selectedProvider) {
        document.getElementById('imapHost').value = selectedProvider.imapHost;
        document.getElementById('imapPort').value = selectedProvider.imapPort;
        
        // Reset SMTP settings
        document.getElementById('enableSmtp').checked = false;
        document.getElementById('smtpSettings').classList.add('hidden');
        document.getElementById('smtpSameAuth').checked = true;
        document.getElementById('smtpAuthFields').classList.add('hidden');
        
        updateAuthModeUi();

        // Update password help text
        const passwordHelp = document.getElementById('passwordHelp');
        if (selectedProvider.requiresAppPassword) {
            passwordHelp.innerHTML = `<span class="mr-1">ℹ️</span>${selectedProvider.notes || 'This provider requires an app-specific password.'}`;
            if (selectedProvider.helpUrl) {
                passwordHelp.innerHTML += ` <a href="${selectedProvider.helpUrl}" target="_blank" class="text-blue-600 hover:underline">Learn more</a>`;
            }
        } else {
            passwordHelp.textContent = '';
        }
    }
}

function updateAuthModeUi() {
    const oauthSection = document.getElementById('oauthSection');
    const passwordSection = document.getElementById('passwordSection');
    const submitButton = document.querySelector('#accountForm button[type="submit"]');
    const oauthConfigHelp = document.getElementById('oauthConfigHelp');
    const microsoftSignInButton = document.getElementById('microsoftSignInButton');

    const supportsOAuth = selectedProvider?.oauthProvider === 'microsoft';
    const preferOAuth = selectedProvider?.id === 'office365';

    if (supportsOAuth) {
        oauthSection.classList.remove('hidden');
        oauthConfigHelp.classList.toggle('hidden', microsoftOAuthConfig.configured);
        microsoftSignInButton.disabled = !microsoftOAuthConfig.configured;
        microsoftSignInButton.classList.toggle('opacity-50', !microsoftOAuthConfig.configured);
        microsoftSignInButton.classList.toggle('cursor-not-allowed', !microsoftOAuthConfig.configured);
    } else {
        oauthSection.classList.add('hidden');
    }

    const showPassword = !supportsOAuth || usePasswordAuth || !preferOAuth;
    passwordSection.classList.toggle('hidden', supportsOAuth && preferOAuth && !usePasswordAuth);
    submitButton.classList.toggle('hidden', supportsOAuth && preferOAuth && !usePasswordAuth);
}

function togglePasswordAuth() {
    usePasswordAuth = true;
    updateAuthModeUi();
}

function startMicrosoftOAuth() {
    const accountName = document.getElementById('accountName').value.trim();
    const email = document.getElementById('email').value.trim();

    if (!accountName || !email) {
        document.getElementById('inlineTestResult').classList.remove('hidden');
        document.getElementById('inlineTestSuccess').classList.add('hidden');
        document.getElementById('inlineTestError').classList.remove('hidden');
        document.getElementById('inlineErrorMessage').textContent = '❌ Please enter an account name and email address first.';
        return;
    }

    const params = new URLSearchParams({
        accountName,
        email,
        providerId: selectedProvider?.id || 'office365',
        host: document.getElementById('imapHost').value || selectedProvider?.imapHost || 'outlook.office365.com',
        port: document.getElementById('imapPort').value || selectedProvider?.imapPort || 993,
        tls: selectedProvider?.imapSecurity !== 'STARTTLS',
        saveToSent: document.getElementById('saveToSent').checked,
    });

    window.location.href = `/api/oauth/microsoft/authorize?${params.toString()}`;
}

// Navigation
function goToStep(step) {
    currentStep = step;
    
    // Hide all steps
    document.getElementById('providerSelection').classList.add('hidden');
    document.getElementById('credentialsForm').classList.add('hidden');
    document.getElementById('testConnection').classList.add('hidden');
    document.getElementById('accountsList').classList.add('hidden');
    
    // Update step indicators
    for (let i = 1; i <= 3; i++) {
        const stepEl = document.getElementById(`step${i}`);
        const circle = stepEl.querySelector('.step-circle');
        
        if (i < step) {
            stepEl.classList.remove('step-inactive');
            circle.classList.add('bg-green-600', 'text-white');
            circle.classList.remove('bg-gray-300', 'text-gray-600', 'bg-blue-600');
            circle.innerHTML = '✓';
        } else if (i === step) {
            stepEl.classList.remove('step-inactive');
            circle.classList.add('bg-blue-600', 'text-white');
            circle.classList.remove('bg-gray-300', 'text-gray-600', 'bg-green-600');
            circle.textContent = i;
        } else {
            stepEl.classList.add('step-inactive');
            circle.classList.add('bg-gray-300', 'text-gray-600');
            circle.classList.remove('bg-blue-600', 'text-white', 'bg-green-600');
            circle.textContent = i;
        }
    }
    
    // Show current step
    switch(step) {
        case 1:
            document.getElementById('providerSelection').classList.remove('hidden');
            break;
        case 2:
            document.getElementById('credentialsForm').classList.remove('hidden');
            break;
        case 3:
            document.getElementById('testConnection').classList.remove('hidden');
            break;
    }
}

// Handle account update
async function handleAccountUpdate(e) {
    e.preventDefault();
    
    const imapUsername = document.getElementById('imapUsername').value.trim();
    const accountData = {
        name: document.getElementById('accountName').value,
        email: document.getElementById('email').value,
        password: document.getElementById('password').value,
        host: document.getElementById('imapHost').value,
        port: parseInt(document.getElementById('imapPort').value),
        tls: selectedProvider?.imapSecurity !== 'STARTTLS',
        saveToSent: document.getElementById('saveToSent').checked,
        imapUsername: imapUsername || undefined
    };

    // Only include password if it was changed
    if (!accountData.password) {
        delete accountData.password;
    }

    goToStep(3);
    await updateAndTestAccount(window.editingAccountId, accountData);
}

// Handle account form submission
async function handleAccountSubmit(e) {
    e.preventDefault();

    const password = document.getElementById('password').value;
    if (!password) {
        document.getElementById('inlineTestResult').classList.remove('hidden');
        document.getElementById('inlineTestSuccess').classList.add('hidden');
        document.getElementById('inlineTestError').classList.remove('hidden');
        document.getElementById('inlineErrorMessage').textContent = 'Please enter a password.';
        return;
    }

    const imapUsername = document.getElementById('imapUsername').value.trim();
    const accountData = {
        name: document.getElementById('accountName').value,
        email: document.getElementById('email').value,
        password: document.getElementById('password').value,
        host: document.getElementById('imapHost').value,
        port: parseInt(document.getElementById('imapPort').value),
        tls: selectedProvider?.imapSecurity !== 'STARTTLS',
        saveToSent: document.getElementById('saveToSent').checked,
        imapUsername: imapUsername || undefined
    };

    // Add SMTP configuration if enabled
    if (document.getElementById('enableSmtp').checked) {
        accountData.smtp = {
            host: document.getElementById('smtpHost').value,
            port: parseInt(document.getElementById('smtpPort').value) || 587,
            secure: document.getElementById('smtpSecure').checked
        };
        
        // Add SMTP auth if not using same credentials
        if (!document.getElementById('smtpSameAuth').checked) {
            accountData.smtp.user = document.getElementById('smtpUser').value;
            accountData.smtp.password = document.getElementById('smtpPassword').value;
        }
    }
    
    // Auto-detect provider if not selected
    if (!selectedProvider || selectedProvider.id === 'custom') {
        const domain = accountData.email.split('@')[1];
        const detectedProvider = providers.find(p => 
            p.domains.some(d => domain.endsWith(d))
        );
        if (detectedProvider) {
            accountData.host = detectedProvider.imapHost;
            accountData.port = detectedProvider.imapPort;
            accountData.tls = detectedProvider.imapSecurity !== 'STARTTLS';
        }
    }
    
    goToStep(3);
    await testAndSaveAccount(accountData);
}

// Update and test account
async function updateAndTestAccount(accountId, accountData) {
    // Show loading
    document.getElementById('testProgress').classList.remove('hidden');
    document.getElementById('testSuccess').classList.add('hidden');
    document.getElementById('testError').classList.add('hidden');
    
    try {
        // Test connection with new data
        const testResponse = await fetch('/api/test-connection', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                ...accountData,
                // If no password provided, we need to test with existing one
                password: accountData.password || 'use-existing'
            })
        });
        
        const testResult = await testResponse.json();
        
        if (!testResult.success && accountData.password) {
            throw new Error(testResult.error || 'Connection test failed');
        }
        
        // Update account
        const updateResponse = await fetch(`/api/accounts/${accountId}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(accountData)
        });
        
        const updateResult = await updateResponse.json();
        
        if (!updateResult.success) {
            throw new Error(updateResult.error || 'Failed to update account');
        }
        
        // Show success
        document.getElementById('testProgress').classList.add('hidden');
        document.getElementById('testSuccess').classList.remove('hidden');
        
        // Update success message
        const successMsg = document.querySelector('#testSuccess h3');
        successMsg.textContent = 'Account updated successfully!';
        
        // Reset edit mode
        window.editingAccountId = null;
        document.getElementById('accountForm').onsubmit = handleAccountSubmit;
        
    } catch (error) {
        // Show error
        document.getElementById('testProgress').classList.add('hidden');
        document.getElementById('testError').classList.remove('hidden');
        document.getElementById('errorMessage').textContent = error.message;
    }
}

// Test and save account
async function testAndSaveAccount(accountData) {
    // Show loading
    document.getElementById('testProgress').classList.remove('hidden');
    document.getElementById('testSuccess').classList.add('hidden');
    document.getElementById('testError').classList.add('hidden');
    
    try {
        // Test connection
        const testResponse = await fetch('/api/test-connection', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(accountData)
        });
        
        const testResult = await testResponse.json();
        
        if (!testResult.success) {
            throw new Error(testResult.error || 'Connection test failed');
        }
        
        // Save account
        const saveResponse = await fetch('/api/accounts', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(accountData)
        });
        
        const saveResult = await saveResponse.json();
        
        if (!saveResult.success) {
            throw new Error(saveResult.error || 'Failed to save account');
        }
        
        // Show success
        document.getElementById('testProgress').classList.add('hidden');
        document.getElementById('testSuccess').classList.remove('hidden');
        
    } catch (error) {
        // Show error
        document.getElementById('testProgress').classList.add('hidden');
        document.getElementById('testError').classList.remove('hidden');
        document.getElementById('errorMessage').textContent = error.message;
    }
}

// View accounts
async function viewAccounts(alertMessage, alertType = 'success') {
    document.getElementById('providerSelection').classList.add('hidden');
    document.getElementById('credentialsForm').classList.add('hidden');
    document.getElementById('testConnection').classList.add('hidden');
    document.getElementById('accountsList').classList.remove('hidden');

    const alert = document.getElementById('accountsListAlert');
    if (alert) {
        alert.classList.add('hidden');
    }
    
    try {
        const response = await fetch('/api/accounts');
        const accounts = await response.json();
        
        const table = document.getElementById('accountsTable');
        if (accounts.length === 0) {
            table.innerHTML = '<p class="text-gray-500 text-center py-8">No accounts configured yet.</p>';
        } else {
            table.innerHTML = `
                <table class="w-full">
                    <thead>
                        <tr class="border-b">
                            <th class="text-left pb-2">Name</th>
                            <th class="text-left pb-2">Email</th>
                            <th class="text-left pb-2">Server</th>
                            <th class="text-left pb-2">Auth</th>
                            <th class="text-right pb-2">Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${accounts.map(account => `
                            <tr class="border-b">
                                <td class="py-3">${account.name}</td>
                                <td class="py-3">${account.user}</td>
                                <td class="py-3">${account.host}</td>
                                <td class="py-3">
                                    <span class="px-2 py-1 text-xs rounded-full ${account.authType === 'oauth2' ? 'bg-blue-100 text-blue-800' : 'bg-green-100 text-green-800'}">
                                        ${account.authType === 'oauth2' ? 'Microsoft OAuth' : 'Password'}
                                    </span>
                                </td>
                                <td class="py-3">
                                    <div class="flex flex-wrap justify-end gap-2">
                                        ${isMicrosoftOAuthAccount(account) ? `
                                            <button
                                                onclick="reauthenticateAccount('${account.id}')"
                                                class="px-3 py-1 text-sm bg-[#0078D4] text-white rounded-md hover:bg-[#106EBE]"
                                                title="Sign in with Microsoft again to refresh access"
                                            >
                                                Re-authenticate
                                            </button>
                                        ` : ''}
                                        <button onclick="editAccount('${account.id}')" class="px-3 py-1 text-sm border border-gray-300 rounded-md hover:bg-gray-50">
                                            Edit
                                        </button>
                                        <button onclick="removeAccount('${account.id}')" class="px-3 py-1 text-sm border border-red-200 text-red-700 rounded-md hover:bg-red-50">
                                            Remove
                                        </button>
                                    </div>
                                </td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            `;
        }

        if (alertMessage) {
            showAccountsAlert(alertMessage, alertType);
        }
    } catch (error) {
        console.error('Failed to load accounts:', error);
        showAccountsAlert('Failed to load accounts: ' + error.message, 'error');
    }
}

// Edit account
async function editAccount(accountId) {
    try {
        // Get account details
        const response = await fetch(`/api/accounts/${accountId}`);
        const result = await response.json();
        
        if (!result.success) {
            throw new Error(result.error || 'Failed to load account');
        }
        
        const account = result.account;
        
        // Store editing account ID
        window.editingAccountId = accountId;
        
        // Clear any stale validation messages
        document.getElementById('inlineTestResult').classList.add('hidden');
        document.getElementById('inlineTestSuccess').classList.add('hidden');
        document.getElementById('inlineTestError').classList.add('hidden');

        // Pre-fill form with account data
        document.getElementById('accountName').value = account.name || '';
        document.getElementById('email').value = account.email || account.user || '';
        document.getElementById('password').value = ''; // Don't pre-fill password
        document.getElementById('password').placeholder = 'Leave blank to keep current password';
        document.getElementById('imapHost').value = account.host || '';
        document.getElementById('imapPort').value = account.port || 993;
        document.getElementById('imapUsername').value = account.email ? account.user : '';
        document.getElementById('saveToSent').checked = account.saveToSent !== false;

        // Pre-fill SMTP settings if available
        if (account.smtp) {
            document.getElementById('enableSmtp').checked = true;
            document.getElementById('smtpSettings').classList.remove('hidden');
            document.getElementById('smtpHost').value = account.smtp.host || '';
            document.getElementById('smtpPort').value = account.smtp.port || 587;
            document.getElementById('smtpSecure').checked = account.smtp.secure || false;
            
            // Check if using different auth
            if (account.smtp.user && account.smtp.user !== account.user) {
                document.getElementById('smtpSameAuth').checked = false;
                document.getElementById('smtpAuthFields').classList.remove('hidden');
                document.getElementById('smtpUser').value = account.smtp.user;
                document.getElementById('smtpPassword').value = '';
                document.getElementById('smtpPassword').placeholder = 'Leave blank to keep current password';
            } else {
                document.getElementById('smtpSameAuth').checked = true;
                document.getElementById('smtpAuthFields').classList.add('hidden');
            }
        } else {
            document.getElementById('enableSmtp').checked = false;
            document.getElementById('smtpSettings').classList.add('hidden');
        }
        
        // Try to detect provider
        const domain = account.user.split('@')[1];
        const detectedProvider = providers.find(p => 
            p.domains.some(d => domain.endsWith(d))
        );
        selectedProvider = detectedProvider || providers.find(p => p.id === 'custom');
        
        // Show credentials form
        goToStep(2);
        
        // Update form submit handler for edit mode
        document.getElementById('accountForm').onsubmit = handleAccountUpdate;
        
        // Add a visual indicator that we're editing
        const formTitle = document.querySelector('#credentialsForm h2');
        formTitle.textContent = 'Edit account details';
        
        // Show test button in edit mode
        document.getElementById('testButton').classList.remove('hidden');
        
        // Update submit button text
        const submitButton = document.querySelector('#accountForm button[type="submit"]');
        submitButton.innerHTML = 'Save Changes<span class="ml-2">→</span>';
        
        // Add cancel edit button
        const backButton = document.querySelector('#credentialsForm button[onclick="goToStep(1)"]');
        backButton.innerHTML = '<span class="mr-2">✕</span>Cancel';
        backButton.onclick = () => {
            window.editingAccountId = null;
            document.getElementById('accountForm').onsubmit = handleAccountSubmit;
            viewAccounts();
        };
        
    } catch (error) {
        alert('Failed to load account: ' + error.message);
    }
}

// Remove account
async function removeAccount(accountId) {
    if (!confirm('Are you sure you want to remove this account?')) return;
    
    try {
        await fetch(`/api/accounts/${accountId}`, { method: 'DELETE' });
        viewAccounts(); // Refresh list
    } catch (error) {
        alert('Failed to remove account: ' + error.message);
    }
}

// Test current settings
async function testCurrentSettings() {
    // Hide previous results
    document.getElementById('inlineTestResult').classList.remove('hidden');
    document.getElementById('inlineTestSuccess').classList.add('hidden');
    document.getElementById('inlineTestError').classList.add('hidden');
    
    // Get current form values
    const imapUsername = document.getElementById('imapUsername').value.trim();
    const accountData = {
        name: document.getElementById('accountName').value,
        email: document.getElementById('email').value,
        password: document.getElementById('password').value,
        host: document.getElementById('imapHost').value,
        port: parseInt(document.getElementById('imapPort').value),
        tls: selectedProvider?.imapSecurity !== 'STARTTLS',
        imapUsername: imapUsername || undefined
    };
    
    // If editing and no password provided, we can't test
    if (window.editingAccountId && !accountData.password) {
        document.getElementById('inlineTestError').classList.remove('hidden');
        document.getElementById('inlineErrorMessage').textContent = '❌ Please enter a password to test the connection';
        return;
    }
    
    // Disable test button during test
    const testButton = document.getElementById('testButton');
    const originalText = testButton.innerHTML;
    testButton.disabled = true;
    testButton.innerHTML = '<span class="mr-2">⏳</span>Testing...';
    
    try {
        const response = await fetch('/api/test-connection', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(accountData)
        });
        
        const result = await response.json();
        
        if (result.success) {
            document.getElementById('inlineTestSuccess').classList.remove('hidden');
        } else {
            document.getElementById('inlineTestError').classList.remove('hidden');
            document.getElementById('inlineErrorMessage').textContent = '❌ ' + (result.error || 'Connection failed');
        }
    } catch (error) {
        document.getElementById('inlineTestError').classList.remove('hidden');
        document.getElementById('inlineErrorMessage').textContent = '❌ ' + error.message;
    } finally {
        // Re-enable test button
        testButton.disabled = false;
        testButton.innerHTML = originalText;
    }
}

// UI helpers
function togglePassword() {
    const input = document.getElementById('password');
    const icon = document.getElementById('passwordToggle');
    
    if (input.type === 'password') {
        input.type = 'text';
        icon.textContent = '🙈';
    } else {
        input.type = 'password';
        icon.textContent = '👁️';
    }
}

function toggleAdvanced() {
    const advanced = document.getElementById('advancedSettings');
    advanced.classList.toggle('hidden');
}

function toggleSmtpSettings() {
    const enabled = document.getElementById('enableSmtp').checked;
    const smtpSettings = document.getElementById('smtpSettings');
    
    if (enabled) {
        smtpSettings.classList.remove('hidden');
        // Auto-fill SMTP settings based on IMAP if provider is selected
        if (selectedProvider) {
            const smtpHost = document.getElementById('smtpHost');
            const smtpPort = document.getElementById('smtpPort');
            
            if (!smtpHost.value) {
                // Convert IMAP host to SMTP host
                smtpHost.value = selectedProvider.imapHost.replace('imap.', 'smtp.').replace('imap-', 'smtp-');
            }
            if (!smtpPort.value) {
                smtpPort.value = '587'; // Default SMTP port
            }
        }
    } else {
        smtpSettings.classList.add('hidden');
    }
}

function toggleSmtpAuth() {
    const sameAuth = document.getElementById('smtpSameAuth').checked;
    const authFields = document.getElementById('smtpAuthFields');
    
    if (sameAuth) {
        authFields.classList.add('hidden');
    } else {
        authFields.classList.remove('hidden');
    }
}

async function reauthenticateAccount(accountId) {
    if (!microsoftOAuthConfig.configured) {
        showAccountsAlert(
            'Microsoft OAuth is not configured. Add MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET to your .env file, then restart npm run setup.',
            'error'
        );
        return;
    }

    window.location.href = `/api/oauth/microsoft/reauth/${accountId}`;
}

function showProviderSelection() {
    // Reset form and edit mode
    document.getElementById('accountForm').reset();
    document.getElementById('password').placeholder = '';
    window.editingAccountId = null;
    document.getElementById('accountForm').onsubmit = handleAccountSubmit;
    selectedProvider = null;
    
    // Reset form title
    const formTitle = document.querySelector('#credentialsForm h2');
    formTitle.textContent = 'Enter your account details';
    
    // Reset back button
    const backButton = document.querySelector('#credentialsForm button[onclick*="goToStep"]');
    if (backButton) {
        backButton.innerHTML = '<span class="mr-2">←</span>Back';
        backButton.onclick = () => goToStep(1);
    }
    
    // Reset submit button text
    const submitButton = document.querySelector('#accountForm button[type="submit"]');
    if (submitButton) {
        submitButton.innerHTML = 'Continue<span class="ml-2">→</span>';
    }
    
    // Hide test button and results
    document.getElementById('testButton').classList.add('hidden');
    document.getElementById('inlineTestResult').classList.add('hidden');
    document.getElementById('inlineTestSuccess').classList.add('hidden');
    document.getElementById('inlineTestError').classList.add('hidden');
    
    currentStep = 1;
    goToStep(1);
}

function addAnotherAccount() {
    // Reset form
    document.getElementById('accountForm').reset();
    selectedProvider = null;
    goToStep(1);
}

function closeWindow() {
    if (window.close) {
        window.close();
    } else {
        alert('Account added successfully! You can close this window.');
    }
}