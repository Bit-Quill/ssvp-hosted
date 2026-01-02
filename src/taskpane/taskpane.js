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

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    bindEvents();
    checkExistingSession();
  } else {
    showError('This add-in requires Excel');
  }
});

function bindEvents() {
  document.getElementById('btn-connect').addEventListener('click', handleConnect);
  document.getElementById('btn-disconnect').addEventListener('click', handleDisconnect);
  document.getElementById('btn-update-context').addEventListener('click', handleUpdateContext);
  document.getElementById('error-close').addEventListener('click', () => {
    document.getElementById('error-message').style.display = 'none';
  });
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

function showManualAuthInstructions(authUrl) {
  hideLoading();

  const errorDiv = document.getElementById('error-message');
  const errorText = document.getElementById('error-text');

  errorText.innerHTML = `
    <strong>🔐 Complete Authentication in Your Browser</strong><br><br>
    <p style="margin-bottom: 12px;">Click the button below to open Snowflake login in your default browser:</p>
    <button id="open-browser-btn" class="primary-button" style="width: 100%; margin-bottom: 8px;">
      Open in Browser
    </button>
    <br>
    <small style="opacity: 0.8;">After you sign in, the window will close automatically and you'll be connected.</small>
    <br><br>
    <details style="margin-top: 8px;">
      <summary style="cursor: pointer; opacity: 0.7;">Or copy URL manually</summary>
      <textarea readonly style="width: 100%; height: 60px; margin: 8px 0; padding: 8px; font-size: 10px; font-family: monospace;">${authUrl}</textarea>
    </details>
  `;
  errorDiv.style.display = 'flex';
  errorDiv.style.background = '#2196F3';

  document.getElementById('open-browser-btn').addEventListener('click', () => {
    window.open(authUrl, '_blank');
  });
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
  errorDiv.style.background = '#dc3545';
  errorDiv.style.color = 'white';

  setTimeout(() => {
    errorDiv.style.display = 'none';
    errorDiv.style.background = '';
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

async function handleUpdateContext() {
  const newWarehouse = document.getElementById('warehouse-select').value;
  const newRole = document.getElementById('role-select').value;

  if (!newWarehouse && !newRole) {
    showError('Please select a warehouse and/or role');
    return;
  }

  try {
    const warehouseChanging = newWarehouse && newWarehouse !== state.warehouse;
    const roleChanging = newRole && newRole !== state.role;

    if (!warehouseChanging && !roleChanging) {
      showError('No changes detected. Select a different warehouse or role.');
      return;
    }

    const changes = [];
    if (warehouseChanging) changes.push(`warehouse to ${newWarehouse}`);
    if (roleChanging) changes.push(`role to ${newRole}`);
    showLoading(`Updating ${changes.join(' and ')}...`);

    SnowflakeAuth.setContext(
      warehouseChanging ? newWarehouse : null,
      roleChanging ? newRole : null
    );

    const result = await SnowflakeAuth.executeStatement(
      'SELECT CURRENT_WAREHOUSE() AS WH, CURRENT_ROLE() AS ROLE'
    );

    const verifiedWarehouse = result.data[0][0];
    const verifiedRole = result.data[0][1];

    if (warehouseChanging && verifiedWarehouse !== newWarehouse) {
      SnowflakeAuth.setContext(state.warehouse, state.role);
      throw new Error(`Warehouse change failed. Expected ${newWarehouse}, got ${verifiedWarehouse}`);
    }

    if (roleChanging && verifiedRole !== newRole) {
      SnowflakeAuth.setContext(state.warehouse, state.role);
      throw new Error(`Role change failed. Expected ${newRole}, got ${verifiedRole}`);
    }

    if (warehouseChanging) state.warehouse = newWarehouse;
    if (roleChanging) state.role = newRole;

    document.getElementById('connected-role').textContent = state.role || '-';
    document.getElementById('current-warehouse').textContent = state.warehouse || '-';
    document.getElementById('current-role').textContent = state.role || '-';
    document.getElementById('current-context').style.display = 'block';

    hideLoading();
    showSuccess(`✅ Updated ${changes.join(' and ')} successfully!`);

    if (roleChanging) {
      await loadWarehousesAndRoles();
    }

  } catch (error) {
    hideLoading();
    showError(`Failed to update context: ${error.message}`);

    if (state.warehouse) {
      document.getElementById('warehouse-select').value = state.warehouse;
    }
    if (state.role) {
      document.getElementById('role-select').value = state.role;
    }
  }
}

window.SnowflakeAuth = SnowflakeAuth;
window.PKCEHelper = PKCEHelper;
window.appState = state;
