import SnowflakeAuth from '../services/snowflake-auth.js';
import PKCEHelper from '../utils/pkce-helper.js';

const state = {
  connected: false,
  account: null,
  warehouse: null,
  role: null,
  availableWarehouses: [],
  availableRoles: []
};

function debugLog(message, data = null) {
  const debugOutput = document.getElementById('debug-output');
  if (debugOutput) {
    const timestamp = new Date().toLocaleTimeString();
    const logEntry = `[${timestamp}] ${message}${data ? '\n' + JSON.stringify(data, null, 2) : ''}\n`;
    debugOutput.textContent += logEntry;
    debugOutput.scrollTop = debugOutput.scrollHeight;
  }
}

Office.onReady((info) => {
  debugLog('Office.js initialized', { host: info.host, platform: info.platform });

  if (info.host === Office.HostType.Excel) {
    initializeApp();
  } else {
    showError('This add-in requires Excel');
  }
});

function initializeApp() {
  bindEvents();
  checkExistingSession();
}

function bindEvents() {
  document.getElementById('btn-connect').addEventListener('click', handleConnect);
  document.getElementById('btn-disconnect').addEventListener('click', handleDisconnect);
  document.getElementById('btn-set-context').addEventListener('click', handleSetContext);
  document.getElementById('btn-change-role').addEventListener('click', handleChangeRole);
  document.getElementById('role-select').addEventListener('change', handleRoleSelectChange);

  document.getElementById('error-close').addEventListener('click', () => {
    document.getElementById('error-message').style.display = 'none';
  });

  document.getElementById('clear-debug').addEventListener('click', () => {
    document.getElementById('debug-output').textContent = '';
  });

  debugLog('Event listeners bound');
}

function checkExistingSession() {
  if (SnowflakeAuth.isAuthenticated()) {
    state.connected = true;
    state.account = SnowflakeAuth.config.account;
    showConnectedState();
  }
}

async function handleConnect() {
  const account = document.getElementById('account-input').value.trim();
  const clientId = document.getElementById('client-id-input').value.trim() || null;
  const role = document.getElementById('role-input').value.trim() || null;

  if (!account) {
    showError('Please enter your Snowflake account identifier');
    return;
  }

  debugLog('Starting connection', { account, clientId, role });

  try {
    showLoading('Initializing authentication...');

    const initConfig = { account };

    if (clientId) {
      initConfig.clientId = clientId;
      initConfig.integrationName = 'Custom OAuth Integration';
    }

    if (role) {
      initConfig.role = role;
    }

    SnowflakeAuth.init(initConfig);

    showLoading('Opening Snowflake login...');
    const result = await SnowflakeAuth.loginWithOAuth();

    state.connected = true;
    state.account = account;

    hideLoading();
    showConnectedState();
    await loadWarehousesAndRoles();

  } catch (error) {
    debugLog('Login failed', { error: error.message });
    hideLoading();

    if (error.message.includes('popup') || error.message.includes('authorization window')) {
      showPopupBlockedMessage();
    } else {
      showError(`Authentication failed: ${error.message}`);
    }
  }
}

async function showPopupBlockedMessage() {
  try {
    const codeVerifier = PKCEHelper.generateCodeVerifier();
    const codeChallenge = await PKCEHelper.generateCodeChallenge(codeVerifier);
    const state = PKCEHelper.generateState();

    PKCEHelper.storePKCEParams(codeVerifier, state);

    const account = document.getElementById('account-input').value.trim();
    const clientId = document.getElementById('client-id-input').value.trim() || null;

    const initConfig = { account };
    if (clientId) {
      initConfig.clientId = clientId;
      initConfig.integrationName = 'Custom OAuth Integration';
    }

    SnowflakeAuth.init(initConfig);
    const authUrl = SnowflakeAuth.buildOAuthUrl(codeChallenge, state);

    const errorDiv = document.getElementById('error-message');
    const errorText = document.getElementById('error-text');

    errorText.innerHTML = `
      <strong>Popup Blocked - Manual Mode</strong><br><br>
      <strong>Copy this URL and open it in your browser:</strong><br>
      <textarea id="manual-auth-url" readonly style="width: 100%; height: 80px; margin: 8px 0; padding: 8px; font-size: 11px; font-family: monospace;">${authUrl}</textarea>
      <button id="copy-url-btn" class="primary-button" style="margin: 8px 0;">
        Copy URL
      </button>
      <button id="open-url-btn" class="secondary-button" style="margin: 8px 0;">
        Open in Browser
      </button>
      <br><br>
      <small>After you approve in your browser, the window will close automatically and you'll be connected.</small>
      <br><br>
      <button id="retry-connect" class="secondary-button" style="margin-top: 8px;">
        Or Retry Popup
      </button>
    `;
    errorDiv.style.display = 'flex';

    document.getElementById('copy-url-btn').addEventListener('click', () => {
      const textarea = document.getElementById('manual-auth-url');
      textarea.select();
      document.execCommand('copy');
      showSuccess('URL copied! Paste it in your browser.');
    });

    document.getElementById('open-url-btn').addEventListener('click', () => {
      const authWindow = window.open(authUrl, 'snowflake-oauth', 'width=500,height=600');

      if (!authWindow) {
        showError('Still blocked. Please copy the URL and paste it in Safari manually.');
      } else {
        errorDiv.style.display = 'none';

        const messageHandler = (event) => {
          if (event.origin === window.location.origin &&
              (event.data.type === 'SNOWFLAKE_AUTH_SUCCESS' || event.data.type === 'SNOWFLAKE_AUTH_ERROR')) {
            window.removeEventListener('message', messageHandler);

            if (authWindow && !authWindow.closed) {
              authWindow.close();
            }

            if (event.data.type === 'SNOWFLAKE_AUTH_SUCCESS') {
              handleAuthCallback(event.data.code, event.data.state);
            } else {
              showError(`Authentication failed: ${event.data.error}`);
            }
          }
        };

        window.addEventListener('message', messageHandler);
      }
    });

    document.getElementById('retry-connect').addEventListener('click', () => {
      errorDiv.style.display = 'none';
      handleConnect();
    });

  } catch (error) {
    showError(`Failed to generate authentication URL: ${error.message}`);
  }
}

async function handleAuthCallback(code, state) {
  try {
    showLoading('Completing authentication...');

    const codeVerifier = PKCEHelper.retrieveCodeVerifier();
    if (!codeVerifier) {
      throw new Error('No PKCE code verifier found');
    }

    if (!PKCEHelper.validateState(state)) {
      throw new Error('Invalid state parameter - possible CSRF attack');
    }

    const tokens = await SnowflakeAuth.exchangeCodeForTokens(code, codeVerifier);

    SnowflakeAuth.accessToken = tokens.access_token;
    SnowflakeAuth.refreshToken = tokens.refresh_token;
    SnowflakeAuth.tokenExpiry = Date.now() + (tokens.expires_in * 1000);
    SnowflakeAuth.authMethod = 'oauth';

    PKCEHelper.clearPKCEParams();

    state.connected = true;
    state.account = SnowflakeAuth.config.account;

    hideLoading();
    showConnectedState();
    await loadWarehousesAndRoles();

  } catch (error) {
    hideLoading();
    showError(`Authentication failed: ${error.message}`);
  }
}

function showSuccess(message) {
  const errorDiv = document.getElementById('error-message');
  const errorText = document.getElementById('error-text');

  errorText.textContent = message;
  errorDiv.style.display = 'flex';
  errorDiv.style.background = '#4CAF50';

  setTimeout(() => {
    errorDiv.style.display = 'none';
    errorDiv.style.background = '';
  }, 3000);
}

function handleDisconnect() {
  try {
    SnowflakeAuth.logout();
    state.connected = false;
    state.account = null;
    showLoginState();
  } catch (error) {
    showError(`Disconnect failed: ${error.message}`);
  }
}

function showLoginState() {
  document.getElementById('login-section').style.display = 'block';
  document.getElementById('connected-section').style.display = 'none';

  const statusDot = document.getElementById('status-dot');
  const statusText = document.getElementById('status-text');
  statusDot.className = 'status-dot disconnected';
  statusText.textContent = 'Not connected';
}

function showConnectedState() {
  document.getElementById('login-section').style.display = 'none';
  document.getElementById('connected-section').style.display = 'block';

  const statusDot = document.getElementById('status-dot');
  const statusText = document.getElementById('status-text');
  statusDot.className = 'status-dot connected';
  statusText.textContent = 'Connected';

  document.getElementById('connected-account').textContent = state.account || '-';
}

function showLoading(message) {
  const overlay = document.getElementById('loading-overlay');
  const loadingText = document.getElementById('loading-text');
  loadingText.textContent = message;
  overlay.style.display = 'flex';
}

function hideLoading() {
  document.getElementById('loading-overlay').style.display = 'none';
}

function showError(message) {
  const errorDiv = document.getElementById('error-message');
  const errorText = document.getElementById('error-text');

  errorText.textContent = message;
  errorDiv.style.display = 'flex';

  setTimeout(() => {
    errorDiv.style.display = 'none';
  }, 10000);
}

async function loadWarehousesAndRoles() {
  try {
    const warehousesResult = await SnowflakeAuth.executeStatement('SHOW WAREHOUSES');
    state.availableWarehouses = warehousesResult.data.map(row => row[0]);

    const rolesResult = await SnowflakeAuth.executeStatement('SHOW ROLES');
    state.availableRoles = rolesResult.data.map(row => row[1]);

    const roleResult = await SnowflakeAuth.executeStatement('SELECT CURRENT_ROLE() AS ROLE');
    const currentRole = roleResult.data[0][0];
    state.role = currentRole;

    populateWarehouseDropdown();
    populateRoleDropdown(currentRole);
    displayCurrentRole(currentRole);

  } catch (error) {
    showError('Failed to load warehouses and roles. You may not have permission.');
  }
}

function populateWarehouseDropdown() {
  const select = document.getElementById('warehouse-select');
  select.innerHTML = '<option value="">Select warehouse...</option>';

  state.availableWarehouses.forEach(warehouse => {
    const option = document.createElement('option');
    option.value = warehouse;
    option.textContent = warehouse;
    select.appendChild(option);
  });
}

function populateRoleDropdown(currentRole) {
  const select = document.getElementById('role-select');
  select.innerHTML = '';

  state.availableRoles.forEach(role => {
    const option = document.createElement('option');
    option.value = role;
    option.textContent = role;

    if (role === currentRole) {
      option.selected = true;
    }

    select.appendChild(option);
  });
}

function displayCurrentRole(role) {
  const roleElement = document.getElementById('connected-role');
  roleElement.textContent = role || '-';
}

async function handleSetContext() {
  const warehouse = document.getElementById('warehouse-select').value;

  if (!warehouse) {
    showError('Please select a warehouse');
    return;
  }

  try {
    showLoading('Setting warehouse...');

    SnowflakeAuth.setContext(warehouse, null);

    const testResult = await SnowflakeAuth.executeStatement(
      'SELECT CURRENT_WAREHOUSE() AS WH, CURRENT_ROLE() AS ROLE'
    );

    state.warehouse = warehouse;

    document.getElementById('current-warehouse').textContent = warehouse;
    document.getElementById('current-role').textContent = state.role || '-';
    document.getElementById('current-context').style.display = 'block';

    hideLoading();

  } catch (error) {
    hideLoading();
    showError(`Failed to set warehouse: ${error.message}`);
  }
}

function handleRoleSelectChange() {
  const selectedRole = document.getElementById('role-select').value;
  const currentRole = state.role;
  const changeRoleButton = document.getElementById('btn-change-role');

  if (selectedRole && selectedRole !== currentRole) {
    changeRoleButton.style.display = 'block';
  } else {
    changeRoleButton.style.display = 'none';
  }
}

async function handleChangeRole() {
  const newRole = document.getElementById('role-select').value;

  if (!newRole) {
    showError('Please select a role');
    return;
  }

  try {
    showLoading('Re-authenticating with new role...');

    const account = state.account;
    const clientId = SnowflakeAuth.config.clientId;

    SnowflakeAuth.logout();

    const initConfig = {
      account: account,
      role: newRole
    };

    if (clientId && clientId !== 'LOCAL_APPLICATION') {
      initConfig.clientId = clientId;
      initConfig.integrationName = 'Custom OAuth Integration';
    }

    SnowflakeAuth.init(initConfig);

    showLoading('Opening Snowflake login...');

    const result = await SnowflakeAuth.loginWithOAuth();

    state.connected = true;
    state.account = account;
    state.role = newRole;

    hideLoading();
    await loadWarehousesAndRoles();

  } catch (error) {
    hideLoading();

    if (error.message.includes('popup') || error.message.includes('authorization window')) {
      showPopupBlockedMessage();
    } else {
      showError(`Failed to change role: ${error.message}`);
    }

    document.getElementById('role-select').value = state.role;
    document.getElementById('btn-change-role').style.display = 'none';
  }
}

window.SnowflakeAuth = SnowflakeAuth;
window.PKCEHelper = PKCEHelper;
window.appState = state;
