import SnowflakeAuth from '../services/snowflake-auth.js';
import PKCEHelper from '../utils/pkce-helper.js';
import SnowflakeAPI from '../services/snowflake-api.js';
import AdhocMetricsManager from '../ui/calculated-fields.js';

const state = {
  connected: false,
  account: null,
  warehouse: null,
  role: null,
  availableWarehouses: [],
  availableRoles: [],
  databases: [],
  schemas: [],
  semanticViews: [],
  selectedDatabase: null,
  selectedSchema: null,
  selectedView: null,
  viewMetadata: null,
  selectedFields: {
    dimensions: [],
    metrics: [],
    facts: []
  },
  // Drag-and-drop zones for pivot table layout
  dropZones: {
    columns: [],  // Fields in columns zone
    rows: [],     // Fields in rows zone
    values: []    // Fields in values zone
  },
  // Track allowed dimensions per metric (for validation)
  allowedDimensionsCache: {},
  // Track if we're in metrics mode or facts mode
  queryMode: null,  // null, 'metrics', or 'facts'
  // Track the last created pivot table for updates
  lastPivotTable: {
    pivotTableName: null,
    pivotSheetName: null,
    dataSheetName: null,
    tableName: null,
    position: null
  }
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Set window title (removes "Excel-" prefix)
    document.title = 'Snowflake Excel Plugin';

    bindEvents();
    checkExistingSession();
  } else {
    showError('This add-in requires Excel');
  }
});

function bindEvents() {
  document.getElementById('btn-connect').addEventListener('click', handleConnect);
  document.getElementById('btn-disconnect').addEventListener('click', handleDisconnect);
  document.getElementById('error-close').addEventListener('click', () => {
    document.getElementById('error-message').classList.add('hidden');
  });

  const roleSelect = document.getElementById('role-select');
  if (roleSelect) {
    roleSelect.addEventListener('change', handleRoleChange);
  }

  const warehouseSelect = document.getElementById('warehouse-select');
  if (warehouseSelect) {
    warehouseSelect.addEventListener('change', () => {
      const btnConnect = document.getElementById('btn-connect-warehouse');
      if (btnConnect) {
        btnConnect.disabled = !warehouseSelect.value || !roleSelect.value;
      }
    });
  }

  const btnConnectWarehouse = document.getElementById('btn-connect-warehouse');
  if (btnConnectWarehouse) {
    btnConnectWarehouse.addEventListener('click', handleConnectWarehouse);
  }

  document.getElementById('btn-change-context').addEventListener('click', showChangeContextModal);
  document.getElementById('modal-close-btn').addEventListener('click', closeChangeContextModal);
  document.getElementById('modal-cancel-btn').addEventListener('click', closeChangeContextModal);
  document.getElementById('modal-apply-btn').addEventListener('click', applyContextChange);

  document.getElementById('database-select').addEventListener('change', handleDatabaseSelect);
  document.getElementById('schema-select').addEventListener('change', handleSchemaSelect);

  const btnCreateAdhocMetric = document.getElementById('btn-create-adhoc-metric');
  if (btnCreateAdhocMetric) {
    btnCreateAdhocMetric.addEventListener('click', () => {
      AdhocMetricsManager.showCreateDialog();
    });
  }

  const btnBackToCatalog = document.getElementById('btn-back-to-catalog');
  if (btnBackToCatalog) {
    btnBackToCatalog.addEventListener('click', handleBackToCatalog);
  }

  const btnDisconnectPivot = document.getElementById('btn-disconnect-pivot');
  if (btnDisconnectPivot) {
    btnDisconnectPivot.addEventListener('click', handleDisconnect);
  }

  const btnCreatePivot = document.getElementById('btn-create-pivot');
  if (btnCreatePivot) {
    btnCreatePivot.addEventListener('click', handleCreatePivot);
  }

  const btnUpdatePivot = document.getElementById('btn-update-pivot');
  if (btnUpdatePivot) {
    btnUpdatePivot.addEventListener('click', handleUpdatePivot);
  }

  const btnShowFilterDemo = document.getElementById('btn-show-filter-demo');
  if (btnShowFilterDemo) {
    btnShowFilterDemo.addEventListener('click', showFilterDemo);
  }

  const btnPreviewQuery = document.getElementById('btn-preview-query');
  if (btnPreviewQuery) {
    btnPreviewQuery.addEventListener('click', handlePreviewQuery);
  }

  // Field list panel toggle
  const fieldListPanelToggle = document.getElementById('field-list-panel-toggle');
  if (fieldListPanelToggle) {
    fieldListPanelToggle.addEventListener('click', toggleFieldListPanel);
  }

  // Pivot table checkbox toggle
  const enablePivotCheckbox = document.getElementById('enable-pivot-table');
  if (enablePivotCheckbox) {
    enablePivotCheckbox.addEventListener('change', handlePivotTableToggle);
  }

  // Filters panel toggle
  const filtersPanelToggle = document.getElementById('filters-panel-toggle');
  if (filtersPanelToggle) {
    filtersPanelToggle.addEventListener('click', toggleFiltersPanel);
  }

  // Query builder panel toggle
  const queryBuilderPanelToggle = document.getElementById('query-builder-panel-toggle');
  if (queryBuilderPanelToggle) {
    queryBuilderPanelToggle.addEventListener('click', toggleQueryBuilderPanel);
  }
}

function handlePivotTableToggle(event) {
  const isChecked = event.target.checked;
  const queryBuilderPanel = document.getElementById('query-builder-panel');
  const queryBuilderContent = document.getElementById('query-builder-panel-content');

  if (isChecked) {
    // Show the query builder panel and content
    queryBuilderPanel.classList.remove('hidden');
    if (queryBuilderContent) {
      queryBuilderContent.classList.remove('hidden');
    }
  } else {
    // Keep the panel visible but hide only the content (drop zones)
    // The panel header with checkbox stays visible so user can re-enable
    if (queryBuilderContent) {
      queryBuilderContent.classList.add('hidden');
    }
  }

  // Save preference
  localStorage.setItem('enablePivotTable', isChecked);
}

function toggleFieldListPanel() {
  const fieldListPanel = document.querySelector('.field-list-panel');
  fieldListPanel.classList.toggle('collapsed');

  // Save state
  const isCollapsed = fieldListPanel.classList.contains('collapsed');
  localStorage.setItem('fieldListPanelCollapsed', isCollapsed);
}

function toggleFiltersPanel() {
  const filtersPanel = document.querySelector('.filters-panel');
  filtersPanel.classList.toggle('collapsed');

  // Save state
  const isCollapsed = filtersPanel.classList.contains('collapsed');
  localStorage.setItem('filtersPanelCollapsed', isCollapsed);
}

function toggleQueryBuilderPanel() {
  const queryBuilderPanel = document.querySelector('.query-builder-panel');
  queryBuilderPanel.classList.toggle('collapsed');

  // Save state
  const isCollapsed = queryBuilderPanel.classList.contains('collapsed');
  localStorage.setItem('queryBuilderPanelCollapsed', isCollapsed);
}

function restoreFieldListPanelState() {
  const isCollapsed = localStorage.getItem('fieldListPanelCollapsed') === 'true';
  const fieldListPanel = document.querySelector('.field-list-panel');

  if (fieldListPanel && isCollapsed) {
    fieldListPanel.classList.add('collapsed');
  }
}

function restoreFiltersPanelState() {
  const isCollapsed = localStorage.getItem('filtersPanelCollapsed') === 'true';
  const filtersPanel = document.querySelector('.filters-panel');

  if (filtersPanel && isCollapsed) {
    filtersPanel.classList.add('collapsed');
  }
}

function restoreQueryBuilderPanelState() {
  const savedState = localStorage.getItem('queryBuilderPanelCollapsed');
  const queryBuilderPanel = document.querySelector('.query-builder-panel');

  if (queryBuilderPanel) {
    // Default to collapsed if no saved state exists
    const isCollapsed = savedState === null ? true : savedState === 'true';

    if (isCollapsed) {
      queryBuilderPanel.classList.add('collapsed');
    }
  }
}

function restorePivotTableCheckboxState() {
  const enablePivotCheckbox = document.getElementById('enable-pivot-table');
  const savedState = localStorage.getItem('enablePivotTable');
  const isEnabled = savedState === 'true';

  if (enablePivotCheckbox) {
    enablePivotCheckbox.checked = isEnabled;

    // Trigger the toggle to update UI
    const queryBuilderPanel = document.getElementById('query-builder-panel');
    const queryBuilderContent = document.getElementById('query-builder-panel-content');
    const createButton = document.getElementById('btn-create-pivot');
    const updateButton = document.getElementById('btn-update-pivot');

    if (isEnabled) {
      queryBuilderPanel?.classList.remove('hidden');
      queryBuilderContent?.classList.remove('hidden');
      if (createButton) {
        createButton.textContent = 'Generate';
      }
      if (updateButton) {
        updateButton.textContent = 'Update';
      }
    } else {
      // Keep panel visible but hide only the content
      queryBuilderPanel?.classList.remove('hidden');
      queryBuilderContent?.classList.add('hidden');
      if (createButton) {
        createButton.textContent = 'Generate';
      }
      if (updateButton) {
        updateButton.textContent = 'Update';
      }
    }
  }
}

async function checkExistingSession() {
  if (SnowflakeAuth.isAuthenticated()) {
    state.connected = true;
    state.account = SnowflakeAuth.config.account;

    try {
      await loadWarehousesAndRoles();
      showRoleSelection();
    } catch (error) {
      showLoginState();
    }
  }
}

async function handleConnect() {
  const account = document.getElementById('account-input').value.trim();

  if (!account) {
    showError('Please enter your Snowflake account identifier');
    return;
  }

  try {
    showLoading('Initializing authentication...');

    const initConfig = { account };
    SnowflakeAuth.init(initConfig);

    // OAuth authentication
    showLoading('Opening Snowflake login...');
    await SnowflakeAuth.loginWithOAuth();

    state.connected = true;
    state.account = account;

    hideLoading();

    await loadWarehousesAndRoles();
    showRoleSelection();

  } catch (error) {
    hideLoading();

    if (error.message.includes('popup') || error.message.includes('authorization window')) {
      showPopupBlockedMessage();
    } else {
      // Log detailed error for debugging
      console.error('Authentication error details:', {
        message: error.message,
        authMethod: SnowflakeAuth.authMethod || 'none',
        hasToken: !!SnowflakeAuth.accessToken,
        stack: error.stack
      });
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

    const clientIdInput = document.getElementById('client-id-input');
    const clientId = clientIdInput ? clientIdInput.value.trim() || null : null;

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
    errorDiv.classList.remove('hidden');

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
        errorDiv.classList.add('hidden');

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
      errorDiv.classList.add('hidden');
      handleConnect();
    });

  } catch (error) {
    showError(`Failed to generate authentication URL: ${error.message}`);
  }
}

async function handleAuthCallback(code, stateParam) {
  try {
    showLoading('Completing authentication...');

    const codeVerifier = PKCEHelper.retrieveCodeVerifier();
    if (!codeVerifier) {
      throw new Error('No PKCE code verifier found');
    }

    if (!PKCEHelper.validateState(stateParam)) {
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

    await loadWarehousesAndRoles();
    showRoleSelection();

  } catch (error) {
    hideLoading();
    showError(`Authentication failed: ${error.message}`);
  }
}

function showSuccess(message) {
  const errorDiv = document.getElementById('error-message');
  const errorText = document.getElementById('error-text');

  errorText.textContent = message;
  errorDiv.classList.remove('hidden');
  errorDiv.classList.add('success');

  setTimeout(() => {
    errorDiv.classList.add('hidden');
    errorDiv.classList.remove('success');
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
  errorDiv.classList.remove('hidden');
  errorDiv.classList.add('info');

  document.getElementById('open-browser-btn').addEventListener('click', () => {
    window.open(authUrl, '_blank');
  });
}

function handleDisconnect() {
  try {
    SnowflakeAuth.logout();

    state.connected = false;
    state.account = null;
    state.warehouse = null;
    state.role = null;
    state.availableWarehouses = [];
    state.availableRoles = [];
    state.databases = [];
    state.schemas = [];
    state.semanticViews = [];
    state.selectedDatabase = null;
    state.selectedSchema = null;
    state.selectedView = null;
    state.viewMetadata = null;

    showLoginState();
  } catch (error) {
    showError(`Disconnect failed: ${error.message}`);
  }
}

function showLoginState() {
  document.getElementById('login-section').classList.remove('hidden');
  document.getElementById('role-warehouse-selection-section').classList.add('hidden');
  document.getElementById('main-explorer-section').classList.add('hidden');

  const statusDot = document.getElementById('status-dot');
  const statusText = document.getElementById('status-text');
  statusDot.className = 'status-dot disconnected';
  statusText.textContent = 'Not connected';
}

function showLoading(message) {
  const overlay = document.getElementById('loading-overlay');
  const loadingText = document.getElementById('loading-text');
  loadingText.textContent = message;
  overlay.classList.remove('hidden');
}

function hideLoading() {
  document.getElementById('loading-overlay').classList.add('hidden');
}

function showError(message) {
  const errorDiv = document.getElementById('error-message');
  const errorText = document.getElementById('error-text');

  errorText.textContent = message;
  errorDiv.classList.remove('hidden', 'success', 'info');
  errorDiv.classList.add('error');

  setTimeout(() => {
    errorDiv.classList.add('hidden');
    errorDiv.classList.remove('error');
  }, 10000);
}

async function loadWarehousesAndRoles() {
  try {
    showLoading('Loading available roles...');

    const rolesResult = await SnowflakeAPI.executeQuery('SHOW ROLES');
    state.availableRoles = rolesResult.rows.map(row => row.name);

    await populateWizardRoleDropdown();
    hideLoading();

  } catch (error) {
    hideLoading();
    throw new Error(`Failed to load roles: ${error.message}`);
  }
}

async function populateWizardRoleDropdown() {
  const select = document.getElementById('role-select');
  select.innerHTML = '';

  state.availableRoles.forEach(role => {
    const option = document.createElement('option');
    option.value = role;
    option.textContent = role;
    select.appendChild(option);
  });

  // Automatically load warehouses for the first role (default selected)
  if (state.availableRoles.length > 0 && select.value) {
    await handleRoleChange({ target: select });
  }
}

function populateWizardWarehouseDropdown() {
  const select = document.getElementById('warehouse-select');
  select.innerHTML = '';

  if (state.availableWarehouses.length === 0) {
    select.innerHTML = '<option value="">No warehouses available for this role</option>';
    return;
  }

  const emptyOption = document.createElement('option');
  emptyOption.value = '';
  emptyOption.textContent = 'Select warehouse...';
  select.appendChild(emptyOption);

  state.availableWarehouses.forEach(warehouse => {
    const option = document.createElement('option');
    option.value = warehouse;
    option.textContent = warehouse;
    select.appendChild(option);
  });
}

async function loadWarehousesForRole(role) {
  try {
    const warehousesResult = await SnowflakeAPI.executeQuery('SHOW WAREHOUSES', { role: role });
    state.availableWarehouses = warehousesResult.rows.map(row => row.name);
    populateWizardWarehouseDropdown();
  } catch (error) {
    throw new Error(`Failed to load warehouses for role ${role}: ${error.message}`);
  }
}

function showRoleSelection() {
  document.getElementById('login-section').classList.add('hidden');
  document.getElementById('role-warehouse-selection-section').classList.remove('hidden');
  document.getElementById('main-explorer-section').classList.add('hidden');

  const statusDot = document.getElementById('status-dot');
  const statusText = document.getElementById('status-text');
  statusDot.className = 'status-dot connected';
  statusText.textContent = 'Select Role & Warehouse';
}

async function handleRoleChange() {
  const selectedRole = document.getElementById('role-select').value;
  const warehouseSelect = document.getElementById('warehouse-select');
  const btnConnect = document.getElementById('btn-connect-warehouse');

  if (!selectedRole) {
    warehouseSelect.innerHTML = '<option value="">Select role first...</option>';
    warehouseSelect.disabled = true;
    if (btnConnect) btnConnect.disabled = true;
    return;
  }

  state.role = selectedRole;

  try {
    warehouseSelect.innerHTML = '<option value="">Loading warehouses...</option>';
    warehouseSelect.disabled = true;
    if (btnConnect) btnConnect.disabled = true;

    // Load warehouses available to this role
    await loadWarehousesForRole(selectedRole);

    warehouseSelect.disabled = false;

    // Check if a warehouse is now selected and enable button if so
    if (warehouseSelect.value && selectedRole) {
      if (btnConnect) btnConnect.disabled = false;
    }

  } catch (error) {
    warehouseSelect.innerHTML = '<option value="">Failed to load warehouses</option>';
    showError(`Failed to load warehouses: ${error.message}`);
  }
}

async function handleConnectWarehouse() {
  const selectedRole = document.getElementById('role-select').value;
  const selectedWarehouse = document.getElementById('warehouse-select').value;

  if (!selectedRole || !selectedWarehouse) {
    showError('Please select both role and warehouse');
    return;
  }

  state.role = selectedRole;
  state.warehouse = selectedWarehouse;

  try {
    showLoading('Setting context...');

    // Set Snowflake context
    SnowflakeAuth.setContext(state.warehouse, state.role);

    // Verify context was set
    const result = await SnowflakeAPI.executeQuery(
      'SELECT CURRENT_WAREHOUSE() AS WH, CURRENT_ROLE() AS ROLE'
    );

    const verifiedWarehouse = result.rows[0].WH;
    const verifiedRole = result.rows[0].ROLE;

    if (verifiedWarehouse !== state.warehouse || verifiedRole !== state.role) {
      throw new Error('Failed to set context. Please try again.');
    }

    hideLoading();

    // Show main explorer
    showMainExplorer();

  } catch (error) {
    hideLoading();
    showError(`Failed to set context: ${error.message}`);
  }
}

async function showMainExplorer() {
  document.getElementById('role-warehouse-selection-section').classList.add('hidden');
  document.getElementById('main-explorer-section').classList.remove('hidden');

  document.getElementById('ctx-role').textContent = state.role;
  document.getElementById('ctx-warehouse').textContent = state.warehouse;

  const statusDot = document.getElementById('status-dot');
  const statusText = document.getElementById('status-text');
  statusDot.className = 'status-dot connected';
  statusText.textContent = 'Connected';

  await loadDatabases();
}

async function showChangeContextModal() {
  const modal = document.getElementById('change-context-modal');
  modal.classList.remove('hidden');

  // Populate dropdowns with current values
  const roleSelect = document.getElementById('modal-role-select');
  const warehouseSelect = document.getElementById('modal-warehouse-select');

  roleSelect.innerHTML = '';
  warehouseSelect.innerHTML = '<option value="">Select role first...</option>';

  // Load available roles
  state.availableRoles.forEach(role => {
    const option = document.createElement('option');
    option.value = role;
    option.textContent = role;
    if (role === state.role) option.selected = true;
    roleSelect.appendChild(option);
  });

  // Add event listener for role change to update warehouses
  roleSelect.addEventListener('change', async () => {
    const selectedRole = roleSelect.value;
    if (!selectedRole) return;

    try {
      warehouseSelect.innerHTML = '<option value="">Loading warehouses...</option>';
      warehouseSelect.disabled = true;

      // Load warehouses for selected role
      await loadWarehousesForRole(selectedRole);

      // Populate modal warehouse dropdown
      warehouseSelect.innerHTML = '';
      warehouseSelect.disabled = false;

      if (state.availableWarehouses.length === 0) {
        warehouseSelect.innerHTML = '<option value="">No warehouses available</option>';
        return;
      }

      state.availableWarehouses.forEach(warehouse => {
        const option = document.createElement('option');
        option.value = warehouse;
        option.textContent = warehouse;
        // Select current warehouse if it's in the list
        if (warehouse === state.warehouse && selectedRole === state.role) {
          option.selected = true;
        }
        warehouseSelect.appendChild(option);
      });

    } catch (error) {
      warehouseSelect.innerHTML = '<option value="">Failed to load warehouses</option>';
      showError(`Failed to load warehouses: ${error.message}`);
    }
  });

  // Trigger initial warehouse load if a role is selected
  if (state.role) {
    roleSelect.dispatchEvent(new Event('change'));
  }
}

function closeChangeContextModal() {
  const modal = document.getElementById('change-context-modal');
  modal.classList.add('hidden');

  const roleSelect = document.getElementById('modal-role-select');
  const newRoleSelect = roleSelect.cloneNode(true);
  roleSelect.parentNode.replaceChild(newRoleSelect, roleSelect);
}

async function applyContextChange() {
  const newRole = document.getElementById('modal-role-select').value;
  const newWarehouse = document.getElementById('modal-warehouse-select').value;

  if (!newRole || !newWarehouse) {
    showError('Please select both role and warehouse');
    return;
  }

  try {
    closeChangeContextModal();
    showLoading('Updating context...');

    // Set new context
    SnowflakeAuth.setContext(newWarehouse, newRole);

    const result = await SnowflakeAPI.executeQuery(
      'SELECT CURRENT_WAREHOUSE() AS WH, CURRENT_ROLE() AS ROLE'
    );

    const verifiedWarehouse = result.rows[0].WH;
    const verifiedRole = result.rows[0].ROLE;

    if (verifiedWarehouse !== newWarehouse || verifiedRole !== newRole) {
      throw new Error('Failed to change context');
    }

    state.warehouse = newWarehouse;
    state.role = newRole;

    document.getElementById('ctx-role').textContent = state.role;
    document.getElementById('ctx-warehouse').textContent = state.warehouse;

    state.databases = [];
    state.schemas = [];
    state.semanticViews = [];
    state.selectedDatabase = null;
    state.selectedSchema = null;
    state.selectedView = null;
    state.viewMetadata = null;

    document.getElementById('database-select').innerHTML = '<option value="">Select database...</option>';
    document.getElementById('schema-select').innerHTML = '<option value="">Select schema...</option>';
    document.getElementById('schema-select').disabled = true;
    document.getElementById('semantic-views-catalog').classList.add('hidden');
    document.getElementById('view-metadata-section').classList.add('hidden');

    await loadDatabases();

    hideLoading();
    showSuccess('Context updated successfully!');

  } catch (error) {
    hideLoading();
    showError(`Failed to update context: ${error.message}`);
  }
}

async function loadDatabases() {
  try {
    showLoading('Loading databases...');
    const databases = await SnowflakeAPI.getDatabases();
    state.databases = databases;

    const select = document.getElementById('database-select');
    select.innerHTML = '<option value="">Select database...</option>';
    databases.forEach(db => {
      const option = document.createElement('option');
      option.value = db.name;
      option.textContent = db.name;
      select.appendChild(option);
    });

    hideLoading();
  } catch (error) {
    hideLoading();
    showError(`Failed to load databases: ${error.message}`);
  }
}

async function handleDatabaseSelect(event) {
  const database = event.target.value;
  state.selectedDatabase = database;

  state.selectedSchema = null;
  state.selectedView = null;
  state.viewMetadata = null;
  document.getElementById('semantic-views-catalog').classList.add('hidden');
  document.getElementById('view-metadata-section').classList.add('hidden');

  if (!database) {
    document.getElementById('schema-select').disabled = true;
    document.getElementById('schema-select').innerHTML = '<option value="">Select schema...</option>';
    return;
  }

  try {
    showLoading('Loading schemas...');
    const schemas = await SnowflakeAPI.getSchemas(database);
    state.schemas = schemas;

    const select = document.getElementById('schema-select');
    select.innerHTML = '<option value="">Select schema...</option>';
    select.disabled = false;

    schemas.forEach(schema => {
      const option = document.createElement('option');
      option.value = schema.name;
      option.textContent = schema.name;
      select.appendChild(option);
    });

    hideLoading();
  } catch (error) {
    hideLoading();
    showError(`Failed to load schemas: ${error.message}`);
  }
}

async function handleSchemaSelect(event) {
  const schema = event.target.value;
  state.selectedSchema = schema;

  state.selectedView = null;
  state.viewMetadata = null;
  document.getElementById('view-metadata-section').classList.add('hidden');

  if (!schema) {
    document.getElementById('semantic-views-catalog').classList.add('hidden');
    return;
  }

  await loadSemanticViews(state.selectedDatabase, schema);
}

async function loadSemanticViews(database, schema) {
  try {
    showLoading('Loading semantic views...');
    const views = await SnowflakeAPI.getSemanticViews(database, schema);
    state.semanticViews = views;

    const catalogSection = document.getElementById('semantic-views-catalog');
    const viewsList = document.getElementById('semantic-views-list');

    if (views.length === 0) {
      viewsList.innerHTML = `
        <div class="empty-state">
          <p>No semantic views found in this schema.</p>
          <small>Make sure you have the necessary privileges.</small>
        </div>
      `;
      catalogSection.classList.remove('hidden');
      hideLoading();
      return;
    }

    viewsList.innerHTML = `
      <ul class="semantic-views-list">
        ${views.map(view => `
          <li class="view-item" data-view-name="${view.name}">
            <div class="view-info">
              <strong>${view.name}</strong>
              ${view.comment ? `<small>${view.comment}</small>` : ''}
            </div>
            <button class="view-action-btn" data-view-name="${view.name}">
              Explore →
            </button>
          </li>
        `).join('')}
      </ul>
    `;

    viewsList.querySelectorAll('.view-action-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const viewName = btn.dataset.viewName;
        handleViewSelect(viewName);
      });
    });

    catalogSection.classList.remove('hidden');
    hideLoading();
  } catch (error) {
    hideLoading();
    showError(`Failed to load semantic views: ${error.message}`);
  }
}

async function handleViewSelect(viewName) {
  const fullyQualifiedName = `${state.selectedDatabase}.${state.selectedSchema}.${viewName}`;
  state.selectedView = fullyQualifiedName;

  // Switch from catalog view to pivot configuration view
  showPivotConfigurationPanel();

  await loadViewMetadata(fullyQualifiedName);
}

/**
 * Switch to Pivot Configuration Panel (hide catalog browser)
 */
function showPivotConfigurationPanel() {
  document.getElementById('main-explorer-section').classList.add('hidden');
  document.getElementById('pivot-configuration-section').classList.remove('hidden');
}

/**
 * Switch back to Catalog Browser Panel (hide pivot configuration)
 */
function showCatalogBrowserPanel() {
  document.getElementById('pivot-configuration-section').classList.add('hidden');
  document.getElementById('main-explorer-section').classList.remove('hidden');
}

/**
 * Handle "Back to Catalog" button click
 */
function handleBackToCatalog() {
  // Clear selection
  state.selectedView = null;
  state.viewMetadata = null;
  state.selectedFields = {
    dimensions: [],
    metrics: [],
    facts: []
  };

  // Clear adhoc metrics
  AdhocMetricsManager.clearAdhocMetrics();

  // Switch back to catalog view
  showCatalogBrowserPanel();
}

async function loadViewMetadata(fullyQualifiedName) {
  try {
    showLoading('Loading view metadata...');

    // Describe the semantic view to get metadata
    const metadata = await SnowflakeAPI.describeSemanticView(fullyQualifiedName);
    state.viewMetadata = metadata;

    // Debug: Show filter details if any exist
    if (metadata.nonEnforcedFilters && metadata.nonEnforcedFilters.length > 0) {
      metadata.nonEnforcedFilters.forEach((filter, i) => {
      });
    } else {
    }

    // Switch to pivot configuration panel
    showPivotConfigurationPanel();

    // View name and debug panel removed for cleaner UI

    // Display unified field list
    displayUnifiedFieldList(metadata);

    // Setup search functionality
    setupFieldSearch();

    // Setup drag and drop
    setupDragAndDrop();

    // Display recommended filters if any
    displayRecommendedFilters(metadata.nonEnforcedFilters || []);

    // Setup add filter button
    setupFilterHandlers();

    // Make the metadata section visible
    document.getElementById('view-metadata-section').classList.remove('hidden');

    // Restore panel collapsed states
    restoreFieldListPanelState();
    restoreFiltersPanelState();
    restoreQueryBuilderPanelState();
    restorePivotTableCheckboxState();

    // Reset button states for new view - enable create, disable update
    const btnCreatePivot = document.getElementById('btn-create-pivot');
    const btnUpdatePivot = document.getElementById('btn-update-pivot');

    if (btnCreatePivot) {
      btnCreatePivot.disabled = false;
    }

    if (btnUpdatePivot) {
      btnUpdatePivot.disabled = true;
    }

    // Clear previous table state since we're loading a new view
    state.lastPivotTable = null;

    // Check if there's an existing pivot table we can update
    await findLastSnowflakePivot();

    hideLoading();
  } catch (error) {
    hideLoading();
    showError(`Failed to load view metadata: ${error.message}`);
    console.error('=== DEBUG: Error loading metadata ===', error);
  }
}

function displayFieldList(containerId, fields, fieldType) {
  const container = document.getElementById(containerId);

  if (!container) {
    console.error(`Container not found: ${containerId}`);
    return;
  }


  if (fields.length === 0) {
    container.innerHTML = '<li class="empty-item">No fields available</li>';
    return;
  }

  container.innerHTML = fields.map(field => {
    const isSelected = state.selectedFields[`${fieldType}s`]?.includes(field.name);
    return `
      <li class="field-item ${isSelected ? 'selected' : ''}" data-field-name="${field.name}" data-field-type="${fieldType}">
        <div class="field-checkbox-wrapper">
          <input
            type="checkbox"
            id="field-${fieldType}-${field.name}"
            class="field-checkbox"
            data-field-name="${field.name}"
            data-field-type="${fieldType}"
            ${isSelected ? 'checked' : ''}
          />
          <label for="field-${fieldType}-${field.name}" class="field-label-wrapper">
            <div class="field-info">
              <span class="field-name">${field.name}</span>
              <span class="field-type">${field.dataType}</span>
            </div>
            ${field.description ? `
              <div class="field-description">
                <small>${field.description}</small>
              </div>
            ` : ''}
            ${field.expression ? `
              <div class="field-expression">
                <code>${field.expression}</code>
              </div>
            ` : ''}
          </label>
        </div>
      </li>
    `;
  }).join('');

  // Bind checkbox change events
  container.querySelectorAll('.field-checkbox').forEach(checkbox => {
    checkbox.addEventListener('change', handleFieldSelection);
  });
}

/**
 * Handle field selection checkbox changes
 */
function handleFieldSelection(event) {
  const checkbox = event.target;
  const fieldName = checkbox.dataset.fieldName;
  const fieldType = checkbox.dataset.fieldType;
  const fieldArrayKey = `${fieldType}s`;

  if (checkbox.checked) {
    // Add field to selected fields
    if (!state.selectedFields[fieldArrayKey].includes(fieldName)) {
      state.selectedFields[fieldArrayKey].push(fieldName);
    }
    // Add selected class to parent field-item
    checkbox.closest('.field-item').classList.add('selected');
  } else {
    // Remove field from selected fields
    state.selectedFields[fieldArrayKey] = state.selectedFields[fieldArrayKey].filter(f => f !== fieldName);
    // Remove selected class
    checkbox.closest('.field-item').classList.remove('selected');
  }

  updateSelectionSummary();
}

/**
 * Update selection summary display
 */
function updateSelectionSummary() {
  const totalSelected =
    state.selectedFields.dimensions.length +
    state.selectedFields.metrics.length +
    state.selectedFields.facts.length;

  // Update button text or add a summary
  const loadButton = document.getElementById('btn-load-to-excel');
  if (loadButton) {
    if (totalSelected > 0) {
      loadButton.textContent = `Load ${totalSelected} Field${totalSelected !== 1 ? 's' : ''} to Excel`;
    } else {
      loadButton.textContent = 'Load Data to Excel';
    }
  }
}

/**
 * Select all fields of a specific type
 */
function selectAllFields(fieldType) {
  if (!state.viewMetadata) return;

  const fields = state.viewMetadata[`${fieldType}s`] || [];
  state.selectedFields[`${fieldType}s`] = fields.map(f => f.name);

  // Update checkboxes
  const containerId = `${fieldType}s-list`;
  const container = document.getElementById(containerId);
  if (container) {
    container.querySelectorAll('.field-checkbox').forEach(cb => {
      cb.checked = true;
      cb.closest('.field-item').classList.add('selected');
    });
  }

  updateSelectionSummary();
}

/**
 * Deselect all fields of a specific type
 */
function deselectAllFields(fieldType) {
  state.selectedFields[`${fieldType}s`] = [];

  // Update checkboxes
  const containerId = `${fieldType}s-list`;
  const container = document.getElementById(containerId);
  if (container) {
    container.querySelectorAll('.field-checkbox').forEach(cb => {
      cb.checked = false;
      cb.closest('.field-item').classList.remove('selected');
    });
  }

  updateSelectionSummary();
}

/**
 * Show demo of what recommended filters would look like
 */
function showFilterDemo() {
  // Create example filters based on common use cases
  const exampleFilters = [
    {
      name: 'recent_data',
      expression: 'sale_date >= CURRENT_DATE - 90',
      description: 'Filter to last 90 days for better query performance'
    },
    {
      name: 'active_customers',
      expression: 'customer_status = \'ACTIVE\'',
      description: 'Exclude inactive customers to reduce dataset size'
    },
    {
      name: 'high_value_transactions',
      expression: 'transaction_amount >= 100',
      description: 'Focus on significant transactions'
    }
  ];

  // Display the example filters
  displayRecommendedFilters(exampleFilters);

  // Show the filters section
  document.getElementById('filters-section').classList.remove('hidden');
  document.getElementById('filters-count').textContent = `${exampleFilters.length} (DEMO)`;

  // Hide the demo button section
  document.getElementById('filters-demo-section').classList.add('hidden');

  // Add a note that these are examples
  const container = document.getElementById('recommended-filters-list');
  const demoNote = document.createElement('li');
  demoNote.style.cssText = 'background: #fff3cd; padding: 12px; margin-top: 8px; border-radius: 4px; border: 1px dashed #ffc107;';
  demoNote.innerHTML = `
    <div style="font-size: 12px;">
      <strong>📌 Note:</strong> These are example filters to demonstrate the feature.<br>
      Your actual semantic view doesn't have any recommended filters defined.<br>
      <button type="button" onclick="hideFilterDemo()" class="link-button small" style="margin-top: 8px;">Hide Demo</button>
    </div>
  `;
  container.appendChild(demoNote);
}

/**
 * Hide filter demo and return to demo section
 */
function hideFilterDemo() {
  document.getElementById('filters-section').classList.add('hidden');
  document.getElementById('filters-demo-section').classList.remove('hidden');
}

// Make hideFilterDemo available globally for onclick
window.hideFilterDemo = hideFilterDemo;

// handleCloseView replaced by handleBackToCatalog

async function handleCreatePivot() {
  if (!state.viewMetadata) {
    showError('Please select a semantic view first');
    return;
  }

  // Check if pivot table mode is enabled
  const enablePivotCheckbox = document.getElementById('enable-pivot-table');
  const isPivotMode = enablePivotCheckbox?.checked || false;

  try {
    showLoading('Preparing query...');

    let allFields = [];

    if (isPivotMode) {
      // Get fields from drop zones
      allFields = [
        ...state.dropZones.columns,
        ...state.dropZones.rows,
        ...state.dropZones.values
      ];

      if (allFields.length === 0) {
        showError('Please drag at least one field to Columns, Rows, or Values to load data.');
        return;
      }
    } else {
      // For data table mode, get all selected fields from the unified field list
      const selectedFields = document.querySelectorAll('.unified-field-item input[type="checkbox"]:checked');

      selectedFields.forEach(checkbox => {
        const fieldItem = checkbox.closest('.unified-field-item');
        const fieldName = fieldItem.dataset.fieldName;
        const fieldType = fieldItem.dataset.fieldType;

        allFields.push({
          name: fieldName,
          type: fieldType
        });
      });

      if (allFields.length === 0) {
        showError('Please select at least one field from the field list.');
        return;
      }
    }

    // Categorize fields by type
    const dimensions = allFields.filter(f => f.type === 'dimension').map(f => f.name);
    const metrics = allFields.filter(f => f.type === 'metric').map(f => f.name);
    const facts = allFields.filter(f => f.type === 'fact').map(f => f.name);
    const adhocMetrics = AdhocMetricsManager.getAdhocMetrics();


    const totalFields = allFields.length + adhocMetrics.length;

    if (isPivotMode) {
      showLoading(`Creating pivot table with ${totalFields} field${totalFields !== 1 ? 's' : ''}...`);
    } else {
      showLoading(`Creating data table with ${totalFields} field${totalFields !== 1 ? 's' : ''}...`);
    }

    // Get custom filters
    const customFilters = getCustomFilters();

    const result = await SnowflakeAPI.executeSemanticQuery({
      semanticView: state.selectedView,
      dimensions: dimensions,
      metrics: metrics,
      facts: facts,
      adhocMetrics: adhocMetrics,
      filters: customFilters,
      limit: 10000
    });

    // Get semantic view name for sheet naming
    const viewParts = state.selectedView.split('.');
    const viewName = viewParts[viewParts.length - 1];

    if (isPivotMode) {
      // Create both data sheet and pivot table sheet
      await writeToExcelWithDataAndPivot(result, viewName, {
        columns: state.dropZones.columns,
        rows: state.dropZones.rows,
        values: state.dropZones.values
      });

      showSuccess(`Successfully created data sheet and pivot table!`);
    } else {
      // Write to Excel - create regular data table in a new sheet
      await writeToExcelAsDataTable(result, `${viewName}_data`);

      showSuccess(`Successfully created data table with ${totalFields} field${totalFields !== 1 ? 's' : ''}!`);
    }

    // Enable the update button and disable the create button now that we have a data table
    const btnCreatePivot = document.getElementById('btn-create-pivot');
    const btnUpdatePivot = document.getElementById('btn-update-pivot');

    if (btnCreatePivot) {
      btnCreatePivot.disabled = true;
    }

    if (btnUpdatePivot) {
      btnUpdatePivot.disabled = false;
    }

    hideLoading();
  } catch (error) {
    hideLoading();

    // Show detailed error with SQL if available
    let errorMessage = error.message;

    // Check if it's a SQL syntax error
    if (errorMessage.includes('SQL compilation error') || errorMessage.includes('syntax error')) {
      console.error('=== SQL SYNTAX ERROR ===');

      // Special handling for semantic view "same entity" error
      if (errorMessage.includes('same entity') || errorMessage.includes('FACTS and DIMENSIONS')) {
        errorMessage = `❌ Snowflake Semantic View Constraint

When using FACTS with DIMENSIONS, they must come from the same logical table.

💡 Solution: Use METRICS instead of FACTS:
• Metrics can aggregate data across joined tables
• Facts are raw values from a single table

Try this:
1. Remove the FACTS from your selection
2. Use only DIMENSIONS + METRICS
3. Or create an Adhoc Metric to aggregate the fact

Example: Instead of using "sale_amount" as a fact, use "SUM(sale_amount)" as a metric.`;
      } else {
        errorMessage = `SQL Syntax Error: ${errorMessage}\n\nℹ️ The generated SQL query has been logged. Check your browser dev tools console (F12) to see the exact query.`;
      }
    }

    showError(errorMessage);
    console.error('=== ERROR creating data table ===', error);
  }
}

async function findLastSnowflakePivot() {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      const pivotTables = workbook.pivotTables;
      pivotTables.load('items');
      await context.sync();

      // Find the most recent SnowflakePivot
      let latestPivot = null;
      let latestTimestamp = 0;

      for (let i = 0; i < pivotTables.items.length; i++) {
        const pivot = pivotTables.items[i];
        pivot.load('name, worksheet');
        await context.sync();

        // Check if it's a Snowflake pivot (name starts with SnowflakePivot_)
        if (pivot.name.startsWith('SnowflakePivot_')) {
          // Extract timestamp from name (SnowflakePivot_123456)
          const match = pivot.name.match(/SnowflakePivot_(\d+)/);
          if (match) {
            const timestamp = parseInt(match[1], 10);
            if (timestamp > latestTimestamp) {
              latestTimestamp = timestamp;
              latestPivot = pivot;
            }
          }
        }
      }

      if (latestPivot) {
        // Found a pivot, now find the corresponding data sheet and table
        latestPivot.load('worksheet');
        await context.sync();

        const pivotSheetName = latestPivot.worksheet.name;
        const dataSheetName = `Data_${latestTimestamp}`;
        const tableName = `SnowflakeData_${latestTimestamp}`;

        // Verify the data sheet exists
        const dataSheet = workbook.worksheets.getItemOrNullObject(dataSheetName);
        await context.sync();

        if (dataSheet.isNullObject) {
          console.warn('Found pivot table but data sheet is missing');
          return false;
        }

        // Get the pivot's position (approximate - we'll use A3 as default for shared sheet)
        const position = 'A3'; // We don't track exact position for existing pivots

        // Restore state
        state.lastPivotTable = {
          pivotTableName: latestPivot.name,
          pivotSheetName: pivotSheetName,
          dataSheetName: dataSheetName,
          tableName: tableName,
          position: position
        };

        // Enable the update button
        const btnUpdatePivot = document.getElementById('btn-update-pivot');
        if (btnUpdatePivot) {
          btnUpdatePivot.disabled = false;
        }

        return true;
      }

      return false;
    });
  } catch (error) {
    console.error('Error finding pivot table:', error);
    return false;
  }
}

async function handleUpdatePivot() {

  if (!state.viewMetadata) {
    showError('Please select a semantic view first');
    return;
  }

  // Check if we have data table info in state
  if (!state.lastPivotTable || !state.lastPivotTable.tableName) {
    showError('No data table to update. Please create a data table first.');
    return;
  }


  // Check if pivot table mode is enabled
  const enablePivotCheckbox = document.getElementById('enable-pivot-table');
  const isPivotMode = enablePivotCheckbox?.checked || false;

  // Check if a pivot table currently exists
  const hasPivotTable = !!(state.lastPivotTable?.pivotTableName && state.lastPivotTable?.pivotSheetName);


  try {
    showLoading('Preparing query...');

    let allFields = [];

    if (isPivotMode) {
      // Get fields from drop zones
      allFields = [
        ...state.dropZones.columns,
        ...state.dropZones.rows,
        ...state.dropZones.values
      ];

      if (allFields.length === 0) {
        showError('Please drag at least one field to Columns, Rows, or Values to update data.');
        return;
      }
    } else {
      // For data table mode, get all selected fields from the unified field list
      const selectedFields = document.querySelectorAll('.unified-field-item input[type="checkbox"]:checked');

      selectedFields.forEach(checkbox => {
        const fieldItem = checkbox.closest('.unified-field-item');
        const fieldName = fieldItem.dataset.fieldName;
        const fieldType = fieldItem.dataset.fieldType;

        allFields.push({
          name: fieldName,
          type: fieldType
        });
      });

      if (allFields.length === 0) {
        showError('Please select at least one field from the field list.');
        return;
      }
    }

    // Categorize fields by type
    const dimensions = allFields.filter(f => f.type === 'dimension').map(f => f.name);
    const metrics = allFields.filter(f => f.type === 'metric').map(f => f.name);
    const facts = allFields.filter(f => f.type === 'fact').map(f => f.name);
    const adhocMetrics = AdhocMetricsManager.getAdhocMetrics();


    const totalFields = allFields.length + adhocMetrics.length;

    showLoading(`Updating data table with ${totalFields} field${totalFields !== 1 ? 's' : ''}...`);

    // Get custom filters
    const customFilters = getCustomFilters();

    const result = await SnowflakeAPI.executeSemanticQuery({
      semanticView: state.selectedView,
      dimensions: dimensions,
      metrics: metrics,
      facts: facts,
      adhocMetrics: adhocMetrics,
      filters: customFilters,
      limit: 10000
    });

    // Update the data table
    await updateExistingDataTable(result);

    if (isPivotMode) {
      // Pivot mode is enabled
      if (hasPivotTable) {
        // Update existing pivot table
        showLoading('Updating pivot table...');
        await updateExistingPivotTable();
        showSuccess(`Successfully updated data table and pivot with ${totalFields} field${totalFields !== 1 ? 's' : ''}!`);
      } else {
        // Create new pivot table
        showLoading('Creating pivot table...');
        await createPivotTableFromExistingData(state.dropZones);
        showSuccess(`Successfully updated data table and created pivot with ${totalFields} field${totalFields !== 1 ? 's' : ''}!`);
      }
    } else {
      // Pivot mode is disabled - just update data table (ignore pivot)
      showSuccess(`Successfully updated data table with ${totalFields} field${totalFields !== 1 ? 's' : ''}!`);
    }

    hideLoading();
  } catch (error) {
    hideLoading();

    // Show detailed error with SQL if available
    let errorMessage = error.message;

    // Check if it's a SQL syntax error
    if (errorMessage.includes('SQL compilation error') || errorMessage.includes('syntax error')) {
      console.error('=== SQL SYNTAX ERROR ===');

      // Special handling for semantic view "same entity" error
      if (errorMessage.includes('same entity') || errorMessage.includes('FACTS and DIMENSIONS')) {
        errorMessage = `❌ Snowflake Semantic View Constraint

When using FACTS with DIMENSIONS, they must come from the same logical table.

💡 Solution: Use METRICS instead of FACTS:
• Metrics can aggregate data across joined tables
• Facts are raw values from a single table

Try this:
1. Remove the FACTS from your selection
2. Use only DIMENSIONS + METRICS
3. Or create an Adhoc Metric to aggregate the fact

Example: Instead of using "sale_amount" as a fact, use "SUM(sale_amount)" as a metric.`;
      } else {
        errorMessage = `SQL Syntax Error: ${errorMessage}\n\nℹ️ The generated SQL query has been logged. Check your browser dev tools console (F12) to see the exact query.`;
      }
    }

    showError(errorMessage);
    console.error('=== ERROR updating data table ===', error);
  }
}

async function updateExistingDataTable(result) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;

    try {

      // Validate state info
      if (!state.lastPivotTable || !state.lastPivotTable.dataSheetName || !state.lastPivotTable.tableName) {
        throw new Error('No data table information found. Please create a data table first.');
      }


      // Get the data sheet and table
      const dataSheet = workbook.worksheets.getItemOrNullObject(state.lastPivotTable.dataSheetName);
      await context.sync();

      if (dataSheet.isNullObject) {
        // Save sheet name before clearing state
        const sheetName = state.lastPivotTable.dataSheetName;
        // Clear the state since the sheet doesn't exist
        state.lastPivotTable = null;
        throw new Error(`Data sheet "${sheetName}" not found. It may have been deleted. Please create a new data table.`);
      }


      // Get the existing table
      const oldTable = dataSheet.tables.getItemOrNullObject(state.lastPivotTable.tableName);
      await context.sync();

      // Delete the old table if it exists
      if (!oldTable.isNullObject) {
        oldTable.delete();
        await context.sync();
      } else {
      }

      // Clear all data from the sheet
      const usedRange = dataSheet.getUsedRange();
      usedRange.clear();
      await context.sync();

      // Write new headers
      const headers = result.columns.map(col => col.name);

      if (headers.length === 0) {
        throw new Error('No columns in result. Please check your query.');
      }

      // Set header range with edge case handling
      let headerRange;
      if (headers.length === 1) {
        headerRange = dataSheet.getRange('A1');
      } else {
        headerRange = dataSheet.getRange('A1').getResizedRange(0, headers.length - 1);
      }

      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.font.size = 11;
      headerRange.format.fill.color = '#4472C4';
      headerRange.format.font.color = '#FFFFFF';

      await context.sync();

      // Write new data rows with edge case handling
      if (result.rows.length > 0) {
        let dataRange;
        if (result.rows.length === 1 && headers.length === 1) {
          // Single cell
          dataRange = dataSheet.getRange('A2');
        } else if (result.rows.length === 1) {
          // Single row, multiple columns
          dataRange = dataSheet.getRange('A2').getResizedRange(0, headers.length - 1);
        } else if (headers.length === 1) {
          // Single column, multiple rows
          dataRange = dataSheet.getRange('A2').getResizedRange(result.rows.length - 1, 0);
        } else {
          // Multiple rows and columns
          dataRange = dataSheet.getRange('A2').getResizedRange(result.rows.length - 1, headers.length - 1);
        }

        const values = result.rows.map(row => headers.map(h => row[h]));
        dataRange.values = values;
      }

      await context.sync();

      // Create new table
      const tableRange = dataSheet.getUsedRange();
      const table = dataSheet.tables.add(tableRange, true);
      table.name = state.lastPivotTable.tableName;
      table.style = 'TableStyleMedium2';

      await context.sync();

      // Auto-fit columns
      tableRange.format.autofitColumns();

      // Activate the data sheet
      dataSheet.activate();

      await context.sync();

    } catch (error) {
      console.error('Error updating data table:', error);
      throw new Error(`Failed to update data table: ${error.message}. You may need to create a new data table instead.`);
    }
  });
}

async function updateExistingPivotTable() {
  await Excel.run(async (context) => {
    const workbook = context.workbook;

    try {

      // Validate state info
      if (!state.lastPivotTable || !state.lastPivotTable.pivotTableName || !state.lastPivotTable.pivotSheetName) {
        return;
      }

      // Get the pivot sheet
      const pivotSheet = workbook.worksheets.getItemOrNullObject(state.lastPivotTable.pivotSheetName);
      await context.sync();

      if (pivotSheet.isNullObject) {
        console.warn(`Pivot sheet "${state.lastPivotTable.pivotSheetName}" not found`);
        return;
      }

      // Get the old pivot table
      const oldPivotTable = workbook.pivotTables.getItemOrNullObject(state.lastPivotTable.pivotTableName);
      await context.sync();

      // Delete the old pivot table if it exists
      if (!oldPivotTable.isNullObject) {
        oldPivotTable.delete();
        await context.sync();
      }

      // Get the data table
      const dataSheet = workbook.worksheets.getItem(state.lastPivotTable.dataSheetName);
      const table = dataSheet.tables.getItem(state.lastPivotTable.tableName);
      await context.sync();

      // Create new pivot table with updated structure
      const pivotTableName = state.lastPivotTable.pivotTableName;
      const pivotTable = workbook.pivotTables.add(pivotTableName, table, pivotSheet.getRange('A1'));

      await context.sync();

      // Add row fields from current drop zones
      if (state.dropZones.rows && state.dropZones.rows.length > 0) {
        for (const field of state.dropZones.rows) {
          const rowHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field.name));
          await context.sync();
        }
      }

      // Add column fields from current drop zones
      if (state.dropZones.columns && state.dropZones.columns.length > 0) {
        for (const field of state.dropZones.columns) {
          const columnHierarchy = pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(field.name));
          await context.sync();
        }
      }

      // Add value fields from current drop zones
      if (state.dropZones.values && state.dropZones.values.length > 0) {
        for (const field of state.dropZones.values) {
          const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(field.name));

          // Set aggregation based on field type
          if (field.type === 'metric' || field.type === 'fact') {
            dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
          } else {
            dataHierarchy.summarizeBy = Excel.AggregationFunction.count;
          }

          await context.sync();
        }
      }

      await context.sync();

      // Activate the pivot sheet
      pivotSheet.activate();

      await context.sync();

    } catch (error) {
      console.error('Error updating pivot table:', error);
      throw new Error(`Failed to update pivot table: ${error.message}`);
    }
  });
}

async function createPivotTableFromExistingData(dropZones) {
  // Generate unique pivot sheet name before entering Excel.run context
  const pivotSheetName = await generateUniqueSheetName('PivotSheet');
  // Extract number from pivot sheet name (e.g., "PivotSheet1" -> "1")
  const sheetNumber = pivotSheetName.replace('PivotSheet', '');

  await Excel.run(async (context) => {
    const workbook = context.workbook;

    try {

      // Validate state info
      if (!state.lastPivotTable || !state.lastPivotTable.dataSheetName || !state.lastPivotTable.tableName) {
        throw new Error('No data table found. Please create a data table first.');
      }

      // Get the existing data table
      const dataSheet = workbook.worksheets.getItem(state.lastPivotTable.dataSheetName);
      const table = dataSheet.tables.getItem(state.lastPivotTable.tableName);
      await context.sync();

      // Create pivot sheet
      const pivotSheet = workbook.worksheets.add(pivotSheetName);
      await context.sync();

      // Create pivot table
      const pivotTableName = `PivotTable${sheetNumber}`;
      const pivotTable = workbook.pivotTables.add(pivotTableName, table, pivotSheet.getRange('A1'));
      await context.sync();


      // Add row fields from drop zones
      if (dropZones.rows && dropZones.rows.length > 0) {
        for (const field of dropZones.rows) {
          const rowHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field.name));
          await context.sync();
        }
      }

      // Add column fields from drop zones
      if (dropZones.columns && dropZones.columns.length > 0) {
        for (const field of dropZones.columns) {
          const columnHierarchy = pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(field.name));
          await context.sync();
        }
      }

      // Add value fields from drop zones
      if (dropZones.values && dropZones.values.length > 0) {
        for (const field of dropZones.values) {
          const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(field.name));

          // Set aggregation based on field type
          if (field.type === 'metric' || field.type === 'fact') {
            dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
          } else {
            dataHierarchy.summarizeBy = Excel.AggregationFunction.count;
          }

          await context.sync();
        }
      }

      await context.sync();

      // Update state to include pivot table info
      state.lastPivotTable = {
        pivotTableName: pivotTableName,
        pivotSheetName: pivotSheet.name,
        dataSheetName: state.lastPivotTable.dataSheetName,
        tableName: state.lastPivotTable.tableName,
        position: 'A1'
      };


      // Activate the pivot sheet
      pivotSheet.activate();

      await context.sync();

    } catch (error) {
      console.error('Error creating pivot table:', error);
      throw new Error(`Failed to create pivot table: ${error.message}`);
    }
  });
}

async function writeToExcel(result) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Write headers
    const headers = result.columns.map(col => col.name);
    const headerRange = sheet.getRange('A1').getResizedRange(0, headers.length - 1);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;

    // Write data
    if (result.rows.length > 0) {
      const dataRange = sheet.getRange('A2').getResizedRange(result.rows.length - 1, headers.length - 1);
      const values = result.rows.map(row => headers.map(h => row[h]));
      dataRange.values = values;
    }

    await context.sync();
  });
}

async function findNextPivotPosition(context, pivotSheet) {
  // Load the used range to find where existing content ends
  const usedRangeOrNull = pivotSheet.getUsedRangeOrNullObject();
  await context.sync();

  if (usedRangeOrNull.isNullObject) {
    // Sheet is empty, start at A3
    return 'A3';
  }

  // Load properties we need
  usedRangeOrNull.load(['columnIndex', 'columnCount', 'rowIndex', 'rowCount']);
  await context.sync();

  const usedRange = usedRangeOrNull;
  const lastColumn = usedRange.columnIndex + usedRange.columnCount;
  const lastRow = usedRange.rowIndex + usedRange.rowCount;


  // Position new pivot to the right with 2 columns padding
  // If we're beyond column AA (column 26), start a new row below
  const maxColumnsPerRow = 26; // Up to column Z

  if (lastColumn + 2 < maxColumnsPerRow) {
    // Place to the right
    const newColumnLetter = getColumnLetter(lastColumn + 2);
    const position = `${newColumnLetter}3`;
    return position;
  } else {
    // Start a new row below with 3 rows padding
    const newRow = lastRow + 3;
    const position = `A${newRow}`;
    return position;
  }
}

function getColumnLetter(columnIndex) {
  // Convert 0-based column index to Excel column letter (0=A, 1=B, etc.)
  let letter = '';
  let temp = columnIndex;

  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }

  return letter;
}

/**
 * Generate a unique sheet name with sequential numbering
 * @param {string} baseName - Base name for the sheet (e.g., "DataSheet" or "PivotSheet")
 * @returns {Promise<string>} - Unique sheet name with number suffix
 */
async function generateUniqueSheetName(baseName) {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();

    const existingNames = sheets.items.map(s => s.name);
    let counter = 1;
    let proposedName = `${baseName}${counter}`;

    // Find the next available number
    while (existingNames.includes(proposedName)) {
      counter++;
      proposedName = `${baseName}${counter}`;
    }

    return proposedName;
  });
}

async function writeToExcelAsDataTable(result, customSheetName = null) {
  // Generate unique sheet name before entering Excel.run context
  const sheetName = customSheetName || await generateUniqueSheetName('DataSheet');

  await Excel.run(async (context) => {
    const workbook = context.workbook;

    // Step 1: Create a new sheet for the data table
    const dataSheet = workbook.worksheets.add(sheetName);

    // Step 2: Write headers
    const headers = result.columns.map(col => col.name);
    const headerRange = dataSheet.getRange('A1').getResizedRange(0, headers.length - 1);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.font.size = 11;
    headerRange.format.fill.color = '#4472C4';
    headerRange.format.font.color = '#FFFFFF';

    await context.sync();

    // Step 3: Write data rows
    if (result.rows.length > 0) {
      const dataRange = dataSheet.getRange('A2').getResizedRange(result.rows.length - 1, headers.length - 1);
      const values = result.rows.map(row => headers.map(h => row[h]));
      dataRange.values = values;
    }

    await context.sync();

    // Step 4: Create an Excel Table from the data
    const tableRange = dataSheet.getUsedRange();
    const table = dataSheet.tables.add(tableRange, true);
    // Extract number from sheet name (e.g., "DataSheet1" -> "1")
    const sheetNumber = sheetName.replace('DataSheet', '');
    const tableName = `DataTable${sheetNumber}`;
    table.name = tableName;
    table.style = 'TableStyleMedium2';

    await context.sync();

    // Step 5: Auto-fit columns
    tableRange.format.autofitColumns();

    // Step 6: Activate the new sheet
    dataSheet.activate();

    await context.sync();

    // Save table info to state for updates
    state.lastPivotTable = {
      pivotTableName: null, // No pivot table (using regular data table)
      pivotSheetName: null,
      dataSheetName: sheetName,
      tableName: tableName,
      position: null
    };

  });
}

async function writeToExcelWithDataAndPivot(result, viewName, dropZones) {
  // Generate unique sheet names before entering Excel.run context
  const dataSheetName = await generateUniqueSheetName('DataSheet');
  const pivotSheetName = await generateUniqueSheetName('PivotSheet');

  await Excel.run(async (context) => {
    const workbook = context.workbook;

    // Step 1: Create the data sheet
    const dataSheet = workbook.worksheets.add(dataSheetName);

    // Write headers and data to data sheet
    const headers = result.columns.map(col => col.name);
    const headerRange = dataSheet.getRange('A1').getResizedRange(0, headers.length - 1);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.font.size = 11;
    headerRange.format.fill.color = '#4472C4';
    headerRange.format.font.color = '#FFFFFF';

    if (result.rows.length > 0) {
      const dataRange = dataSheet.getRange('A2').getResizedRange(result.rows.length - 1, headers.length - 1);
      const values = result.rows.map(row => headers.map(h => row[h]));
      dataRange.values = values;
    }

    await context.sync();

    // Create an Excel Table from the data
    const tableRange = dataSheet.getUsedRange();
    const table = dataSheet.tables.add(tableRange, true);
    // Extract number from sheet name (e.g., "DataSheet1" -> "1")
    const sheetNumber = dataSheetName.replace('DataSheet', '');
    const tableName = `DataTable${sheetNumber}`;
    table.name = tableName;
    table.style = 'TableStyleMedium2';

    await context.sync();

    // Auto-fit columns
    tableRange.format.autofitColumns();

    await context.sync();

    // Step 2: Create the pivot table sheet
    const pivotSheet = workbook.worksheets.add(pivotSheetName);

    await context.sync();

    // Create the pivot table - source should be the table we just created
    const pivotTableName = `PivotTable${sheetNumber}`;
    const pivotTable = workbook.pivotTables.add(pivotTableName, table, pivotSheet.getRange('A1'));

    await context.sync();

    // Add row fields
    if (dropZones.rows && dropZones.rows.length > 0) {
      for (const field of dropZones.rows) {
        const rowHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field.name));
        await context.sync();
      }
    }

    // Add column fields
    if (dropZones.columns && dropZones.columns.length > 0) {
      for (const field of dropZones.columns) {
        const columnHierarchy = pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(field.name));
        await context.sync();
      }
    }

    // Add value fields
    if (dropZones.values && dropZones.values.length > 0) {
      for (const field of dropZones.values) {
        const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(field.name));

        // Set aggregation based on field type
        if (field.type === 'metric' || field.type === 'fact') {
          dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
        } else {
          dataHierarchy.summarizeBy = Excel.AggregationFunction.count;
        }

        await context.sync();
      }
    }

    await context.sync();

    // Activate the pivot sheet
    pivotSheet.activate();

    // Load sheet names before accessing them
    dataSheet.load('name');
    pivotSheet.load('name');

    await context.sync();

    // Save pivot table info to state
    state.lastPivotTable = {
      pivotTableName: pivotTableName,
      pivotSheetName: pivotSheet.name,
      dataSheetName: dataSheet.name,
      tableName: tableName,
      position: 'A1'
    };

  });
}

async function writeToExcelWithPivot(result) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const timestamp = Date.now().toString().slice(-6);

    // Step 1: Create a new hidden data sheet for this pivot's source data
    const dataSheetName = `Data_${timestamp}`;
    const dataSheet = workbook.worksheets.add(dataSheetName);

    // Write headers and data to data sheet
    const headers = result.columns.map(col => col.name);
    const headerRange = dataSheet.getRange('A1').getResizedRange(0, headers.length - 1);
    headerRange.values = [headers];

    if (result.rows.length > 0) {
      const dataRange = dataSheet.getRange('A2').getResizedRange(result.rows.length - 1, headers.length - 1);
      const values = result.rows.map(row => headers.map(h => row[h]));
      dataRange.values = values;
    }

    await context.sync();

    // Step 2: Create an Excel Table from the data
    const tableRange = dataSheet.getUsedRange();
    const table = dataSheet.tables.add(tableRange, true);
    const tableName = `SnowflakeData_${timestamp}`;
    table.name = tableName;
    table.style = 'TableStyleMedium2';

    await context.sync();

    // Step 3: Get or create the shared "Snowflake Pivots" sheet
    const sharedPivotSheetName = 'Snowflake_Pivots';
    let pivotSheet = workbook.worksheets.getItemOrNullObject(sharedPivotSheetName);
    await context.sync();

    if (pivotSheet.isNullObject) {
      // Create the shared sheet if it doesn't exist
      pivotSheet = workbook.worksheets.add(sharedPivotSheetName);
      await context.sync();
    } else {
    }

    // Step 4: Find the next available position on the sheet
    const position = await findNextPivotPosition(context, pivotSheet);

    // Step 4.5: Add a label above the pivot table
    const labelRow = parseInt(position.match(/\d+/)[0]) - 1; // One row above the pivot
    const labelColumn = position.match(/[A-Z]+/)[0]; // Same column as pivot
    const labelRange = pivotSheet.getRange(`${labelColumn}${labelRow}`);
    const currentTime = new Date().toLocaleTimeString();
    labelRange.values = [[`Pivot Table - ${currentTime}`]];
    labelRange.format.font.bold = true;
    labelRange.format.font.size = 12;
    labelRange.format.font.color = '#1E88E5';
    await context.sync();

    // Step 5: Create PivotTable at the calculated position
    const pivotTableName = `SnowflakePivot_${timestamp}`;
    const pivotTable = workbook.pivotTables.add(
      pivotTableName,
      table,
      pivotSheet.getRange(position)
    );

    await context.sync();

    // Save pivot table info to state for updates
    state.lastPivotTable = {
      pivotTableName: pivotTableName,
      pivotSheetName: sharedPivotSheetName,
      dataSheetName: dataSheetName,
      tableName: tableName,
      position: position
    };

    // Step 5: Configure PivotTable based on drop zones

    // Add fields to ROWS
    if (state.dropZones.rows && state.dropZones.rows.length > 0) {
      for (const field of state.dropZones.rows) {
        try {
          const hierarchy = pivotTable.hierarchies.getItem(field.name);
          pivotTable.rowHierarchies.add(hierarchy);
        } catch (err) {
          console.warn(`Could not add ${field.name} to rows:`, err);
        }
      }
    }

    await context.sync();

    // Add fields to COLUMNS
    if (state.dropZones.columns && state.dropZones.columns.length > 0) {
      for (const field of state.dropZones.columns) {
        try {
          const hierarchy = pivotTable.hierarchies.getItem(field.name);
          pivotTable.columnHierarchies.add(hierarchy);
        } catch (err) {
          console.warn(`Could not add ${field.name} to columns:`, err);
        }
      }
    }

    await context.sync();

    // Add fields to VALUES
    if (state.dropZones.values && state.dropZones.values.length > 0) {
      for (const field of state.dropZones.values) {
        try {
          const hierarchy = pivotTable.hierarchies.getItem(field.name);
          const dataHierarchy = pivotTable.dataHierarchies.add(hierarchy);

          // Set aggregation based on field type
          if (field.type === 'metric') {
            dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
          } else if (field.type === 'fact') {
            // Facts can be summed or counted depending on data type
            dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
          } else {
            // Dimensions should be counted
            dataHierarchy.summarizeBy = Excel.AggregationFunction.count;
          }

        } catch (err) {
          console.warn(`Could not add ${field.name} to values:`, err);
        }
      }
    }

    await context.sync();

    // Step 6: Hide the data sheet and activate pivot sheet
    dataSheet.visibility = Excel.SheetVisibility.hidden;
    pivotSheet.activate();

    await context.sync();

  });
}

async function updateExistingPivot(result) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;

    try {

      // First, verify all components exist
      const dataSheet = workbook.worksheets.getItemOrNullObject(state.lastPivotTable.dataSheetName);
      const pivotSheet = workbook.worksheets.getItemOrNullObject(state.lastPivotTable.pivotSheetName);
      const pivotTable = workbook.pivotTables.getItemOrNullObject(state.lastPivotTable.pivotTableName);

      await context.sync();

      // Check if components exist
      if (dataSheet.isNullObject) {
        throw new Error(`Data sheet "${state.lastPivotTable.dataSheetName}" not found. It may have been deleted.`);
      }
      if (pivotSheet.isNullObject) {
        throw new Error(`Pivot sheet "${state.lastPivotTable.pivotSheetName}" not found. It may have been deleted.`);
      }
      if (pivotTable.isNullObject) {
        throw new Error(`Pivot table "${state.lastPivotTable.pivotTableName}" not found. It may have been deleted.`);
      }


      // Clear existing data in the data sheet
      const usedRange = dataSheet.getUsedRange();
      usedRange.clear();
      await context.sync();

      // Write new headers and data
      const headers = result.columns.map(col => col.name);
      const headerRange = dataSheet.getRange('A1').getResizedRange(0, headers.length - 1);
      headerRange.values = [headers];

      if (result.rows.length > 0) {
        const dataRange = dataSheet.getRange('A2').getResizedRange(result.rows.length - 1, headers.length - 1);
        const values = result.rows.map(row => headers.map(h => row[h]));
        dataRange.values = values;
      }

      await context.sync();

      // Delete the old table and create a new one with the same name
      const oldTable = dataSheet.tables.getItemOrNullObject(state.lastPivotTable.tableName);
      await context.sync();

      if (!oldTable.isNullObject) {
        oldTable.delete();
        await context.sync();
      }

      // Create new table with same name
      const tableRange = dataSheet.getUsedRange();
      const table = dataSheet.tables.add(tableRange, true);
      table.name = state.lastPivotTable.tableName;
      table.style = 'TableStyleMedium2';

      await context.sync();

      // Delete the old pivot table and create a new one at the same position
      pivotTable.delete();
      await context.sync();

      // Note: We don't clear the entire sheet since there may be other pivots on it

      // Create a new pivot table with the same name at the same position
      const newPivotTable = workbook.pivotTables.add(
        state.lastPivotTable.pivotTableName,
        table,
        pivotSheet.getRange(state.lastPivotTable.position || 'A3')
      );
      await context.sync();

      // Configure the new pivot table based on current drop zones

      // Add fields to ROWS
      if (state.dropZones.rows && state.dropZones.rows.length > 0) {
        for (const field of state.dropZones.rows) {
          try {
            const hierarchy = newPivotTable.hierarchies.getItem(field.name);
            newPivotTable.rowHierarchies.add(hierarchy);
          } catch (err) {
            console.warn(`Could not add ${field.name} to rows:`, err);
          }
        }
        await context.sync();
      }

      // Add fields to COLUMNS
      if (state.dropZones.columns && state.dropZones.columns.length > 0) {
        for (const field of state.dropZones.columns) {
          try {
            const hierarchy = newPivotTable.hierarchies.getItem(field.name);
            newPivotTable.columnHierarchies.add(hierarchy);
          } catch (err) {
            console.warn(`Could not add ${field.name} to columns:`, err);
          }
        }
        await context.sync();
      }

      // Add fields to VALUES
      if (state.dropZones.values && state.dropZones.values.length > 0) {
        for (const field of state.dropZones.values) {
          try {
            const hierarchy = newPivotTable.hierarchies.getItem(field.name);
            const dataHierarchy = newPivotTable.dataHierarchies.add(hierarchy);

            // Set aggregation based on field type
            if (field.type === 'metric') {
              dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
            } else if (field.type === 'fact') {
              dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
            } else {
              // Dimensions should be counted
              dataHierarchy.summarizeBy = Excel.AggregationFunction.count;
            }

          } catch (err) {
            console.warn(`Could not add ${field.name} to values:`, err);
          }
        }
        await context.sync();
      }

      // Activate the pivot sheet to show the updated data
      pivotSheet.activate();
      await context.sync();

    } catch (error) {
      console.error('Error updating pivot table:', error);
      throw new Error(`Failed to update pivot table: ${error.message}. You may need to create a new pivot table instead.`);
    }
  });
}

async function handlePreviewQuery() {
  if (!state.viewMetadata) {
    showError('Please select a semantic view first');
    return;
  }

  // Get ACTUAL selected fields (not all fields)
  const dimensions = state.selectedFields.dimensions;
  const metrics = state.selectedFields.metrics;
  const facts = state.selectedFields.facts;
  const adhocMetrics = AdhocMetricsManager.getAdhocMetrics();

  const totalFields = dimensions.length + metrics.length + facts.length + adhocMetrics.length;

  if (totalFields === 0) {
    showError('Please select at least one field first');
    return;
  }

  // Build SELECT clause
  const selectFields = [...dimensions, ...metrics, ...facts];

  // Add adhoc metrics
  adhocMetrics.forEach(adhoc => {
    selectFields.push(`${adhoc.expression} AS ${adhoc.name}`);
  });

  const selectClause = selectFields.join(',\n  ');

  // Build GROUP BY clause
  let groupByClause = '';
  if (metrics.length > 0 || adhocMetrics.length > 0) {
    const groupByFields = [...dimensions, ...facts];
    if (groupByFields.length > 0) {
      groupByClause = `\nGROUP BY ${groupByFields.join(', ')}`;
    }
  }

  // Build final query
  let query = `SELECT\n  ${selectClause}\nFROM ${state.selectedView}${groupByClause}\nLIMIT 10000`;

  // Show in a modal dialog
  showQueryPreviewDialog(query, selectFields.length, adhocMetrics.length);
}

function showQueryPreviewDialog(query, fieldCount, adhocCount) {
  // Create modal overlay
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 9999;';

  overlay.innerHTML = `
    <div style="background: white; padding: 24px; border-radius: 8px; max-width: 600px; max-height: 80vh; overflow: auto; box-shadow: 0 4px 12px rgba(0,0,0,0.3);">
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
        <h3 style="margin: 0; color: #333;">📝 Query Preview</h3>
        <button onclick="this.closest('.modal-overlay').remove()" style="background: none; border: none; font-size: 24px; cursor: pointer; color: #666;">&times;</button>
      </div>

      <div style="background: #f8f9fa; padding: 12px; border-radius: 4px; margin-bottom: 16px;">
        <strong>Selected Fields:</strong> ${fieldCount} total
        ${adhocCount > 0 ? `<br><strong>Adhoc Metrics:</strong> ${adhocCount}` : ''}
      </div>

      <div style="background: #2d2d2d; color: #f8f8f2; padding: 16px; border-radius: 4px; overflow-x: auto; font-family: 'Monaco', 'Menlo', 'Courier New', monospace; font-size: 12px; line-height: 1.5; white-space: pre;">
${query}
      </div>

      <div style="margin-top: 16px; display: flex; gap: 8px; justify-content: flex-end;">
        <button onclick="navigator.clipboard.writeText(\`${query.replace(/`/g, '\\`')}\`); alert('Query copied to clipboard!');" style="padding: 8px 16px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer;">
          Copy Query
        </button>
        <button onclick="this.closest('.modal-overlay').remove()" style="padding: 8px 16px; background: #e0e0e0; border: none; border-radius: 4px; cursor: pointer;">
          Close
        </button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);
}

// === FILTERS FUNCTIONS ===

let customFiltersCount = 0;

function displayRecommendedFilters(filters) {
  const container = document.getElementById('recommended-filters-container');
  const listContainer = document.getElementById('recommended-filters-list');

  if (!filters || filters.length === 0) {
    container.classList.add('hidden');
    return;
  }

  container.classList.remove('hidden');

  listContainer.innerHTML = filters.map(filter => `
    <div class="recommended-filter-item">
      <strong>${filter.name}</strong>
      ${filter.description ? `<div style="font-size: 9px; color: #666; margin-bottom: 4px;">${filter.description}</div>` : ''}
      <div class="filter-expression">${filter.expression}</div>
    </div>
  `).join('');

}

function setupFilterHandlers() {
  const btnAddFilter = document.getElementById('btn-add-filter');
  if (btnAddFilter) {
    btnAddFilter.addEventListener('click', addCustomFilter);
  }
}

function addCustomFilter() {
  const container = document.getElementById('custom-filters-container');
  const filterId = `filter-${customFiltersCount++}`;

  // Get all available fields for the dropdown
  const allFields = [
    ...state.viewMetadata.dimensions.map(f => ({ name: f.name, type: 'dimension' })),
    ...state.viewMetadata.metrics.map(f => ({ name: f.name, type: 'metric' })),
    ...state.viewMetadata.facts.map(f => ({ name: f.name, type: 'fact' }))
  ].sort((a, b) => a.name.localeCompare(b.name));

  const filterRow = document.createElement('div');
  filterRow.className = 'filter-row';
  filterRow.dataset.filterId = filterId;

  filterRow.innerHTML = `
    <div class="filter-row-main">
      <select class="filter-field" data-filter-id="${filterId}">
        <option value="">Select field...</option>
        ${allFields.map(f => `<option value="${f.name}" data-field-type="${f.type}">${f.name}</option>`).join('')}
      </select>

      <select class="filter-operator" data-filter-id="${filterId}">
        <option value="">Select field first...</option>
      </select>

      <button type="button" class="remove-filter-btn" onclick="removeCustomFilter('${filterId}')">×</button>
    </div>
    <div class="filter-values-container">
      <input type="text" class="filter-values-textbox" data-filter-id="${filterId}" readonly placeholder="Click to select values..." title="Click to select values" />
    </div>
  `;

  container.appendChild(filterRow);

  // Add event listener for field selection
  const fieldSelect = filterRow.querySelector('.filter-field');
  const operatorSelect = filterRow.querySelector('.filter-operator');
  const valuesTextbox = filterRow.querySelector('.filter-values-textbox');

  fieldSelect.addEventListener('change', async (e) => {
    const fieldName = e.target.value;
    const selectedOption = e.target.selectedOptions[0];
    const fieldType = selectedOption?.dataset.fieldType;

    // Store field type in the filter row
    filterRow.dataset.fieldType = fieldType || '';

    if (!fieldName) {
      operatorSelect.innerHTML = '<option value="">Select field first...</option>';
      operatorSelect.disabled = false;
      valuesTextbox.disabled = true;
      valuesTextbox.value = '';
      valuesTextbox.placeholder = 'Select field first...';
      filterRow.dataset.selectedValues = '[]';
      return;
    }

    try {
      // Block operator dropdown and show loading state
      operatorSelect.innerHTML = '<option value="">Loading...</option>';
      operatorSelect.disabled = true;
      valuesTextbox.disabled = true;
      valuesTextbox.value = '';
      valuesTextbox.placeholder = 'Loading values...';

      // Query for distinct values (already ordered by SQL)
      // This makes a network call to Snowflake: SELECT DISTINCT field FROM view LIMIT 1000
      const distinctValues = await SnowflakeAPI.getDistinctValues(
        state.selectedView,
        fieldName,
        1000
      );

      // Detect data type from values
      const dataType = detectDataType(distinctValues);
      filterRow.dataset.dataType = dataType;

      // Update operators based on detected data type
      updateOperatorsForDataType(operatorSelect, dataType);

      // Enable operator dropdown now that we know the data type
      operatorSelect.disabled = false;

      // Store all values for filtering and reset selected values
      filterRow.dataset.allValues = JSON.stringify(distinctValues);
      filterRow.dataset.selectedValues = '[]';

      // Enable textbox
      valuesTextbox.disabled = false;
      valuesTextbox.placeholder = 'Click to select values...';
      valuesTextbox.title = 'Click to select values';

    } catch (error) {
      operatorSelect.innerHTML = '<option value="">Error</option>';
      operatorSelect.disabled = true;
      valuesTextbox.disabled = true;
      valuesTextbox.placeholder = 'Error loading values';
      showError(`Failed to load values for ${fieldName}: ${error.message}`);
      console.error('Error loading distinct values:', error);
    }
  });

  // Add click handler for values textbox
  valuesTextbox.addEventListener('click', () => {
    if (!valuesTextbox.disabled) {
      openFilterValueModal(filterId, filterRow);
    }
  });
}

// Detect data type from values
function detectDataType(values) {
  if (!values || values.length === 0) return 'text';

  // Sample first few non-null values
  const sample = values.filter(v => v !== null && v !== undefined && v !== '').slice(0, 20);
  if (sample.length === 0) return 'text';

  let numericCount = 0;
  let dateCount = 0;

  for (const value of sample) {
    const str = String(value).trim();

    // Check if numeric (including decimals)
    if (/^-?\d+(\.\d+)?$/.test(str)) {
      numericCount++;
      continue;
    }

    // Check if date (various formats)
    const datePatterns = [
      /^\d{4}-\d{2}-\d{2}/, // 2024-01-15
      /^\d{2}\/\d{2}\/\d{4}/, // 01/15/2024
      /^\d{4}\/\d{2}\/\d{2}/, // 2024/01/15
    ];

    if (datePatterns.some(pattern => pattern.test(str))) {
      dateCount++;
    }
  }

  // If most values are numeric
  if (numericCount / sample.length > 0.8) return 'numeric';

  // If most values are dates
  if (dateCount / sample.length > 0.8) return 'date';

  // Default to text
  return 'text';
}

// Update operators based on data type
function updateOperatorsForDataType(operatorSelect, dataType) {
  // Preserve current selection if exists
  const currentValue = operatorSelect.value;

  let operators = [];

  if (dataType === 'numeric' || dataType === 'date') {
    // Numeric or date fields: support all comparison operators including BETWEEN
    operators = [
      { value: 'IN', label: 'IN' },
      { value: '=', label: '=' },
      { value: '!=', label: '!=' },
      { value: '>', label: '>' },
      { value: '<', label: '<' },
      { value: '>=', label: '>=' },
      { value: '<=', label: '<=' },
      { value: 'BETWEEN', label: 'BETWEEN' }
    ];
  } else {
    // Text fields: support IN, =, !=, LIKE
    operators = [
      { value: 'IN', label: 'IN' },
      { value: '=', label: '=' },
      { value: '!=', label: '!=' },
      { value: 'LIKE', label: 'LIKE' }
    ];
  }

  operatorSelect.innerHTML = operators.map(op =>
    `<option value="${op.value}">${op.label}</option>`
  ).join('');

  // Restore selection if the operator is still available
  if (currentValue && operators.find(op => op.value === currentValue)) {
    operatorSelect.value = currentValue;
  }
}

// Open modal to select filter values
function openFilterValueModal(filterId, filterRow) {
  const allValues = JSON.parse(filterRow.dataset.allValues || '[]');
  const selectedValues = JSON.parse(filterRow.dataset.selectedValues || '[]');

  // Create modal overlay
  const modalOverlay = document.createElement('div');
  modalOverlay.className = 'filter-modal-overlay';
  modalOverlay.innerHTML = `
    <div class="filter-modal">
      <div class="filter-modal-header">
        <h3>Select Values</h3>
        <button type="button" class="filter-modal-close">×</button>
      </div>
      <div class="filter-modal-body">
        <div class="filter-modal-search">
          <input type="text" class="filter-modal-search-input" placeholder="Search values..." />
          <span class="filter-modal-count">${allValues.length} values</span>
        </div>
        <div class="filter-modal-actions">
          <button type="button" class="filter-modal-select-all">Select All</button>
          <button type="button" class="filter-modal-clear-all">Clear All</button>
        </div>
        <div class="filter-modal-list"></div>
      </div>
      <div class="filter-modal-footer">
        <button type="button" class="filter-modal-cancel">Cancel</button>
        <button type="button" class="filter-modal-apply">Apply</button>
      </div>
    </div>
  `;

  document.body.appendChild(modalOverlay);

  const modalList = modalOverlay.querySelector('.filter-modal-list');
  const searchInput = modalOverlay.querySelector('.filter-modal-search-input');
  const selectAllBtn = modalOverlay.querySelector('.filter-modal-select-all');
  const clearAllBtn = modalOverlay.querySelector('.filter-modal-clear-all');
  const applyBtn = modalOverlay.querySelector('.filter-modal-apply');
  const cancelBtn = modalOverlay.querySelector('.filter-modal-cancel');
  const closeBtn = modalOverlay.querySelector('.filter-modal-close');

  // Render values
  function renderValues(values) {
    modalList.innerHTML = values.slice(0, 200).map((val) => {
      const displayValue = val === null ? '(null)' : val;
      const safeValue = String(val).replace(/"/g, '&quot;');
      const isChecked = selectedValues.includes(val);
      return `
        <label class="filter-modal-item">
          <input type="checkbox" value="${safeValue}" ${isChecked ? 'checked' : ''} />
          <span>${displayValue}</span>
        </label>
      `;
    }).join('');

    if (values.length > 200) {
      const moreDiv = document.createElement('div');
      moreDiv.className = 'filter-modal-more';
      moreDiv.textContent = `+ ${values.length - 200} more (use search)`;
      modalList.appendChild(moreDiv);
    }
  }

  renderValues(allValues);

  // Search functionality
  searchInput.addEventListener('input', (e) => {
    const searchTerm = e.target.value.toLowerCase();
    const filtered = allValues.filter(val => {
      const display = val === null ? '(null)' : String(val).toLowerCase();
      return display.includes(searchTerm);
    });
    renderValues(filtered);
  });

  // Select/Clear all
  selectAllBtn.addEventListener('click', () => {
    modalList.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
  });

  clearAllBtn.addEventListener('click', () => {
    modalList.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
  });

  // Apply
  applyBtn.addEventListener('click', () => {
    const checked = Array.from(modalList.querySelectorAll('input[type="checkbox"]:checked'));
    const values = checked.map(cb => cb.value);

    // Update stored selected values
    filterRow.dataset.selectedValues = JSON.stringify(values);

    // Update textbox with selected values
    const valuesTextbox = filterRow.querySelector('.filter-values-textbox');

    if (values.length > 0) {
      // Show comma-separated values in the textbox
      const displayText = values.join(', ');
      valuesTextbox.value = displayText;
      valuesTextbox.title = `${values.length} value(s) selected: ${displayText}`;
    } else {
      valuesTextbox.value = '';
      valuesTextbox.placeholder = 'Click to select values...';
      valuesTextbox.title = 'Click to select values';
    }

    modalOverlay.remove();
  });

  // Cancel/Close
  const closeModal = () => modalOverlay.remove();
  cancelBtn.addEventListener('click', closeModal);
  closeBtn.addEventListener('click', closeModal);
  modalOverlay.addEventListener('click', (e) => {
    if (e.target === modalOverlay) closeModal();
  });
}

function removeCustomFilter(filterId) {
  const filterRow = document.querySelector(`[data-filter-id="${filterId}"]`);
  if (filterRow) {
    filterRow.remove();
  }
}

function getCustomFilters() {
  const filters = [];
  const filterRows = document.querySelectorAll('.filter-row');

  filterRows.forEach(row => {
    const field = row.querySelector('.filter-field').value;
    const operator = row.querySelector('.filter-operator').value;
    const selectedValues = JSON.parse(row.dataset.selectedValues || '[]');

    if (!field || !operator || selectedValues.length === 0) return;

    if (operator === 'BETWEEN') {
      // BETWEEN requires exactly 2 values (min and max)
      if (selectedValues.length >= 2) {
        // Take first 2 values and sort them
        const values = [selectedValues[0], selectedValues[1]].sort((a, b) => {
          // Try numeric sort first
          const numA = parseFloat(a);
          const numB = parseFloat(b);
          if (!isNaN(numA) && !isNaN(numB)) {
            return numA - numB;
          }
          // Fall back to string sort
          return String(a).localeCompare(String(b));
        });
        filters.push({ field, operator, values });
      }
    } else {
      // Handle regular operators with value selection
      filters.push({ field, operator, values: selectedValues });
    }
  });

  return filters;
}

// === UNIFIED FIELD LIST FUNCTIONS ===

function displayUnifiedFieldList(metadata) {
  const container = document.getElementById('unified-field-list');
  if (!container) {
    console.error('Unified field list container not found');
    return;
  }

  // Combine all fields with their types
  const allFields = [
    ...metadata.dimensions.map(f => ({ ...f, fieldType: 'dimension', icon: 'D' })),
    ...metadata.metrics.map(f => ({ ...f, fieldType: 'metric', icon: 'M' })),
    ...metadata.facts.map(f => ({ ...f, fieldType: 'fact', icon: 'F' }))
  ];

  // Sort alphabetically
  allFields.sort((a, b) => a.name.localeCompare(b.name));


  container.innerHTML = allFields.map(field => `
    <li class="unified-field-item"
        data-field-name="${field.name}"
        data-field-type="${field.fieldType}"
        data-field-icon="${field.icon}"
        data-logical-table="${field.logicalTable || ''}">
      <input type="checkbox"
             class="field-checkbox"
             id="field-${field.fieldType}-${field.name}"
             data-field-name="${field.name}"
             data-field-type="${field.fieldType}"
             data-field-icon="${field.icon}"
             data-logical-table="${field.logicalTable || ''}">
      <label for="field-${field.fieldType}-${field.name}"
             class="field-label"
             draggable="true">
        <div class="field-type-icon ${field.fieldType}">${field.icon}</div>
        <span class="unified-field-name">${field.name}</span>
      </label>
    </li>
  `).join('');

  // Store all fields in state for search
  state.allFields = allFields;

  // Add event listeners for checkboxes
  const checkboxes = container.querySelectorAll('.field-checkbox');
  checkboxes.forEach(checkbox => {
    checkbox.addEventListener('change', handleFieldCheckboxChange);
  });
}

async function handleFieldCheckboxChange(event) {
  const checkbox = event.target;
  const fieldData = {
    name: checkbox.dataset.fieldName,
    type: checkbox.dataset.fieldType,
    icon: checkbox.dataset.fieldIcon,
    logicalTable: checkbox.dataset.logicalTable || null
  };

  if (checkbox.checked) {
    // Add field to appropriate drop zone
    let targetZone = null;

    // Determine default zone based on field type
    if (fieldData.type === 'dimension') {
      targetZone = 'rows'; // Dimensions typically go in rows for grouping
    } else if (fieldData.type === 'metric' || fieldData.type === 'fact') {
      targetZone = 'values'; // Metrics and facts go in values
    }

    if (targetZone) {
      // Validate before adding
      const validation = await validateFieldAddition(fieldData);
      if (!validation.valid) {
        // Uncheck the checkbox if validation fails
        checkbox.checked = false;
        showError(validation.error);
        return;
      }

      // Get the target drop zone element
      const dropZone = document.getElementById(`${targetZone}-drop-zone`);

      // Check if field already exists in this zone
      const existingChip = dropZone.querySelector(`[data-chip-field="${fieldData.name}"]`);
      if (!existingChip) {
        // Create and add the chip
        const chip = createFieldChip(fieldData, targetZone);
        dropZone.appendChild(chip);

        // Update state
        await updateDropZones();

      }
    }
  } else {
    // Remove field from all drop zones
    const allZones = ['columns', 'rows', 'values'];

    allZones.forEach(zoneName => {
      const dropZone = document.getElementById(`${zoneName}-drop-zone`);
      const chip = dropZone.querySelector(`[data-chip-field="${fieldData.name}"]`);
      if (chip) {
        chip.remove();
      }
    });

    // Update state
    await updateDropZones();

  }
}

function setupFieldSearch() {
  const searchInput = document.getElementById('field-search');
  if (!searchInput) return;

  searchInput.addEventListener('input', (e) => {
    const searchTerm = e.target.value.toLowerCase().trim();
    const fieldItems = document.querySelectorAll('.unified-field-item');

    fieldItems.forEach(item => {
      const fieldName = item.dataset.fieldName.toLowerCase();
      if (fieldName.includes(searchTerm)) {
        item.style.display = 'flex';
      } else {
        item.style.display = 'none';
      }
    });
  });
}

function setupDragAndDrop() {
  const fieldList = document.getElementById('unified-field-list');
  const dropZones = document.querySelectorAll('.drop-zone-area');

  // Drag from field list (only from label, not checkbox)
  fieldList?.addEventListener('dragstart', (e) => {
    // Check if we're dragging from the label
    const label = e.target.closest('.field-label');
    const isFromCheckbox = e.target.classList.contains('field-checkbox');

    if (label && !isFromCheckbox) {
      const fieldItem = label.closest('.unified-field-item');

      // Prevent dragging disabled fields
      if (fieldItem && fieldItem.classList.contains('disabled')) {
        e.preventDefault();
        return;
      }

      if (fieldItem) {
        fieldItem.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'copy';
        e.dataTransfer.setData('text/plain', JSON.stringify({
          name: fieldItem.dataset.fieldName,
          type: fieldItem.dataset.fieldType,
          icon: fieldItem.dataset.fieldIcon,
          logicalTable: fieldItem.dataset.logicalTable || null
        }));
      }
    } else if (isFromCheckbox) {
      // Prevent dragging from checkbox
      e.preventDefault();
    }
  });

  fieldList?.addEventListener('dragend', (e) => {
    const label = e.target.closest('.field-label');
    if (label) {
      const fieldItem = label.closest('.unified-field-item');
      if (fieldItem) {
        fieldItem.classList.remove('dragging');
      }
    }
  });

  // Drop zones
  dropZones.forEach(zone => {
    zone.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
      zone.classList.add('drag-over');
    });

    zone.addEventListener('dragleave', (e) => {
      if (e.target === zone) {
        zone.classList.remove('drag-over');
      }
    });

    zone.addEventListener('drop', async (e) => {
      e.preventDefault();
      zone.classList.remove('drag-over');

      try {
        const fieldData = JSON.parse(e.dataTransfer.getData('text/plain'));
        const zoneName = zone.dataset.zone;

        // Check if field already exists in this zone
        const existingChip = zone.querySelector(`[data-chip-field="${fieldData.name}"]`);
        if (existingChip) {
          return;
        }

        // Validate before adding
        const validation = await validateFieldAddition(fieldData);
        if (!validation.valid) {
          showError(validation.error);
          return;
        }

        // Create field chip
        const chip = createFieldChip(fieldData, zoneName);
        zone.appendChild(chip);

        // Check the corresponding checkbox
        const checkbox = document.querySelector(`.field-checkbox[data-field-name="${fieldData.name}"][data-field-type="${fieldData.type}"]`);
        if (checkbox) {
          checkbox.checked = true;
        }

        // Update state
        await updateDropZones();

      } catch (error) {
        showError('Failed to add field: ' + error.message);
      }
    });
  });

  // Also setup drag between zones
  document.addEventListener('dragstart', (e) => {
    if (e.target.classList.contains('field-chip')) {
      e.dataTransfer.effectAllowed = 'move';
      const chipData = {
        name: e.target.dataset.chipField,
        type: e.target.dataset.chipType,
        icon: e.target.dataset.chipIcon
      };
      e.dataTransfer.setData('text/plain', JSON.stringify(chipData));
      e.target.style.opacity = '0.5';
    }
  });

  document.addEventListener('dragend', (e) => {
    if (e.target.classList.contains('field-chip')) {
      e.target.style.opacity = '1';
    }
  });
}

function createFieldChip(fieldData, zoneName) {
  const chip = document.createElement('div');
  chip.className = 'field-chip';
  chip.draggable = true;
  chip.dataset.chipField = fieldData.name;
  chip.dataset.chipType = fieldData.type;
  chip.dataset.chipIcon = fieldData.icon;
  chip.dataset.chipZone = zoneName;
  chip.dataset.logicalTable = fieldData.logicalTable || '';

  chip.innerHTML = `
    <div class="field-chip-icon">${fieldData.icon}</div>
    <span>${fieldData.name}</span>
    <button class="remove-chip" onclick="removeChip(this)">×</button>
  `;

  return chip;
}

function removeChip(button) {
  const chip = button.closest('.field-chip');
  const fieldName = chip.dataset.chipField;
  const fieldType = chip.dataset.chipType;

  // Remove the chip
  chip.remove();

  // Uncheck the corresponding checkbox
  const checkbox = document.querySelector(`.field-checkbox[data-field-name="${fieldName}"][data-field-type="${fieldType}"]`);
  if (checkbox) {
    checkbox.checked = false;
  }

  updateDropZones();
}

async function updateDropZones() {
  // Update state based on current chips in drop zones
  const zones = ['columns', 'rows', 'values'];

  // Clear drop zones state
  state.dropZones = { columns: [], rows: [], values: [] };

  zones.forEach(zone => {
    const zoneElement = document.getElementById(`${zone}-drop-zone`);
    const chips = zoneElement?.querySelectorAll('.field-chip') || [];

    state.dropZones[zone] = Array.from(chips).map(chip => ({
      name: chip.dataset.chipField,
      type: chip.dataset.chipType,
      icon: chip.dataset.chipIcon,
      logicalTable: chip.dataset.logicalTable || null
    }));
  });

  // Determine query mode based on what's in the zones
  const allFields = [
    ...state.dropZones.columns,
    ...state.dropZones.rows,
    ...state.dropZones.values
  ];

  const hasMetrics = allFields.some(f => f.type === 'metric');
  const hasFacts = allFields.some(f => f.type === 'fact');

  if (hasMetrics) {
    state.queryMode = 'metrics';
  } else if (hasFacts) {
    state.queryMode = 'facts';
  } else {
    state.queryMode = null;
  }


  // Show warning if facts are being used
  showFactsWarningIfNeeded();

  // Show info about metric-dimension constraints
  showMetricDimensionInfoIfNeeded();

  // Update UI to show/hide disabled fields
  await updateFieldListDisabledStates();
}

function showFactsWarningIfNeeded() {
  const allFields = [
    ...state.dropZones.columns,
    ...state.dropZones.rows,
    ...state.dropZones.values
  ];

  const facts = allFields.filter(f => f.type === 'fact');
  const dimensions = allFields.filter(f => f.type === 'dimension');

  // Get or create warning element
  let warningEl = document.getElementById('facts-warning');

  if (facts.length > 0 && dimensions.length > 0) {
    // Analyze table compatibility
    const factTables = new Set();
    facts.forEach(f => {
      if (f.logicalTable) {
        factTables.add(f.logicalTable);
      }
    });

    const dimensionTables = new Set();
    dimensions.forEach(d => {
      if (d.logicalTable) {
        dimensionTables.add(d.logicalTable);
      }
    });

    // Check for compatibility
    const hasIncompatibleDimensions = Array.from(dimensionTables).some(dt => !factTables.has(dt));

    let warningMessage = '';
    let warningStyle = '';

    if (factTables.size > 1) {
      // Multiple fact tables - ERROR
      warningStyle = 'background: #f8d7da; border: 1px solid #dc3545;';
      warningMessage = `
        <strong style="color: #721c24;">❌ Error: Multiple Fact Tables</strong>
        <p style="margin: 4px 0; color: #721c24;">Facts are from multiple tables: ${Array.from(factTables).join(', ')}. This is not allowed.</p>
        <p style="margin: 4px 0; color: #721c24;"><strong>Solution:</strong> Keep facts from only one table, or use Metrics instead.</p>
      `;
    } else if (hasIncompatibleDimensions && factTables.size > 0) {
      // Dimensions from other tables - ERROR
      warningStyle = 'background: #f8d7da; border: 1px solid #dc3545;';
      const factTableName = Array.from(factTables)[0];
      const incompatibleDimTables = Array.from(dimensionTables).filter(dt => dt !== factTableName);
      warningMessage = `
        <strong style="color: #721c24;">❌ Error: Incompatible Dimensions</strong>
        <p style="margin: 4px 0; color: #721c24;">Facts are from table '${factTableName}', but some dimensions are from other tables: ${incompatibleDimTables.join(', ')}</p>
        <p style="margin: 4px 0; color: #721c24;"><strong>Solutions:</strong></p>
        <ul style="margin: 4px 0 0 16px; color: #721c24;">
          <li>Remove dimensions from other tables</li>
          <li>Use only dimensions from '${factTableName}'</li>
          <li>Or use Metrics instead of Facts (recommended)</li>
        </ul>
      `;
    }

    // Only show warning if there's an actual error (not just info)
    if (warningMessage) {
      // Show or update warning
      if (!warningEl) {
        warningEl = document.createElement('div');
        warningEl.id = 'facts-warning';
        warningEl.style.cssText = warningStyle + ' border-radius: 4px; padding: 8px; margin: 8px 0; font-size: 11px;';

        // Insert before query actions bar
        const actionsBar = document.querySelector('.query-actions-bar');
        if (actionsBar) {
          actionsBar.parentNode.insertBefore(warningEl, actionsBar);
        }
      } else {
        warningEl.style.cssText = warningStyle + ' border-radius: 4px; padding: 8px; margin: 8px 0; font-size: 11px;';
      }

      warningEl.innerHTML = warningMessage;
    } else {
      // Remove info card if everything is valid
      if (warningEl) {
        warningEl.remove();
      }
    }
  } else {
    // Remove warning if it exists
    if (warningEl) {
      warningEl.remove();
    }
  }
}

function showMetricDimensionInfoIfNeeded() {
  // Info card disabled per user request
  // Remove any existing info element if present
  const infoEl = document.getElementById('metric-dimension-info');
  if (infoEl) {
    infoEl.remove();
  }
}

async function validateFieldAddition(fieldData) {
  const { name, type, logicalTable } = fieldData;

  // Get all currently selected fields
  const allFields = [
    ...state.dropZones.columns,
    ...state.dropZones.rows,
    ...state.dropZones.values
  ];

  const hasMetrics = allFields.some(f => f.type === 'metric');
  const hasFacts = allFields.some(f => f.type === 'fact');
  const facts = allFields.filter(f => f.type === 'fact');

  // Rule 1: Cannot mix metrics and facts
  if (type === 'metric' && hasFacts) {
    return {
      valid: false,
      error: 'Cannot add metrics when facts are already selected. Remove facts first, or use only facts without metrics.'
    };
  }

  if (type === 'fact' && hasMetrics) {
    return {
      valid: false,
      error: 'Cannot add facts when metrics are already selected. Remove metrics first, or use only metrics without facts.'
    };
  }

  // NEW RULE: When facts are selected, enforce same logical table constraint
  if (type === 'fact' && hasFacts) {
    // Get logical tables of already selected facts
    const factTables = new Set();
    facts.forEach(f => {
      const factMeta = state.viewMetadata.facts.find(fm => fm.name === f.name);
      if (factMeta && factMeta.logicalTable) {
        factTables.add(factMeta.logicalTable);
      }
    });

    // Check if new fact is from same table
    if (logicalTable && factTables.size > 0 && !factTables.has(logicalTable)) {
      const tableNames = Array.from(factTables).join(', ');
      return {
        valid: false,
        error: `Facts must come from the same logical table. Selected facts are from '${tableNames}', but "${name}" is from '${logicalTable}'. Remove existing facts or choose a fact from the same table.`
      };
    }
  }

  // NEW RULE: When facts are selected, dimensions must be from same table
  if (type === 'dimension' && hasFacts) {
    // Get logical tables of already selected facts
    const factTables = new Set();
    facts.forEach(f => {
      const factMeta = state.viewMetadata.facts.find(fm => fm.name === f.name);
      if (factMeta && factMeta.logicalTable) {
        factTables.add(factMeta.logicalTable);
      }
    });

    // Check if dimension is from same table as facts
    if (logicalTable && factTables.size > 0 && !factTables.has(logicalTable)) {
      const tableNames = Array.from(factTables).join(', ');
      return {
        valid: false,
        error: `When using facts from '${tableNames}', dimensions must be from the same logical table. "${name}" is from '${logicalTable}'. Either remove the facts or choose a dimension from the facts' table.`
      };
    }
  }

  // Rule 2: If adding a dimension and there are metrics, check compatibility
  // Constraint: Dimension's logical table must be related to metric's logical table
  // and have equal or lower granularity (validated via SHOW SEMANTIC DIMENSIONS FOR METRIC)
  if (type === 'dimension' && hasMetrics) {
    const metrics = allFields.filter(f => f.type === 'metric');

    // Check if this dimension is allowed for all metrics
    for (const metric of metrics) {
      const allowed = await isDimensionAllowedForMetric(name, metric.name);
      if (!allowed) {
        return {
          valid: false,
          error: `Dimension "${name}" cannot be used with metric "${metric.name}".\n\n` +
                 `Reason: The dimension's logical table must be related to the metric's logical table and have equal or lower granularity.\n\n` +
                 `Solutions:\n` +
                 `• Remove metric "${metric.name}"\n` +
                 `• Choose a different dimension that's compatible with this metric\n` +
                 `• Check relationship definitions in your semantic view`
        };
      }
    }
  }

  // Rule 3: If adding a metric and there are dimensions, check compatibility
  // Same constraint applies: metric must be compatible with all selected dimensions
  if (type === 'metric') {
    const dimensions = allFields.filter(f => f.type === 'dimension');

    for (const dimension of dimensions) {
      const allowed = await isDimensionAllowedForMetric(dimension.name, name);
      if (!allowed) {
        return {
          valid: false,
          error: `Metric "${name}" cannot be used with dimension "${dimension.name}".\n\n` +
                 `Reason: The metric's logical table must be related to the dimension's logical table.\n\n` +
                 `Solutions:\n` +
                 `• Remove dimension "${dimension.name}"\n` +
                 `• Choose a different metric that's compatible with your dimensions\n` +
                 `• Check relationship definitions in your semantic view`
        };
      }
    }
  }

  return { valid: true };
}

async function isDimensionAllowedForMetric(dimensionName, metricName) {
  try {
    // Check cache first
    const cacheKey = `${metricName}`;
    if (!state.allowedDimensionsCache[cacheKey]) {
      const fetchMsg = `    🔄 Fetching allowed dimensions for "${metricName}" from Snowflake...`;

      // Fetch allowed dimensions for this metric
      const allowedDims = await SnowflakeAPI.getDimensionsForMetric(
        state.selectedView,
        metricName
      );
      state.allowedDimensionsCache[cacheKey] = allowedDims.map(d => d.name);

      const cacheMsg = `    ✅ Cached ${allowedDims.length} allowed dimensions: ${allowedDims.map(d => d.name).join(', ')}`;
    } else {
    }

    const allowedDimensions = state.allowedDimensionsCache[cacheKey];
    const isAllowed = allowedDimensions.includes(dimensionName);
    const checkResultMsg = `    ${isAllowed ? '✅' : '❌'} "${dimensionName}" ${isAllowed ? 'IS' : 'IS NOT'} in allowed list [${allowedDimensions.join(', ')}]`;

    return isAllowed;
  } catch (error) {
    const errorMsg = `    ⚠️ ERROR checking "${dimensionName}" + "${metricName}": ${error.message}`;

    // If we can't check, allow it (fail open)
    const failOpenMsg = `    ⚠️ FAIL OPEN - allowing "${dimensionName}" with "${metricName}"`;
    return true;
  }
}

async function updateFieldListDisabledStates() {
  const fieldList = document.getElementById('unified-field-list');
  if (!fieldList) {
    return;
  }

  const allItems = fieldList.querySelectorAll('.unified-field-item');

  // Get all currently selected fields
  const allFields = [
    ...state.dropZones.columns,
    ...state.dropZones.rows,
    ...state.dropZones.values
  ];

  const hasMetrics = allFields.some(f => f.type === 'metric');
  const hasFacts = allFields.some(f => f.type === 'fact');
  const hasDimensions = allFields.some(f => f.type === 'dimension');
  const metrics = allFields.filter(f => f.type === 'metric');
  const facts = allFields.filter(f => f.type === 'fact');
  const dimensions = allFields.filter(f => f.type === 'dimension');

  const debugInfo = {
    totalFields: allFields.length,
    hasMetrics,
    hasFacts,
    hasDimensions,
    metricCount: metrics.length,
    factCount: facts.length,
    dimensionCount: dimensions.length,
    metricNames: metrics.map(m => m.name),
    factNames: facts.map(f => f.name),
    dimensionNames: dimensions.map(d => d.name),
    viewMetadataDimensions: state.viewMetadata?.dimensions?.length || 0,
    viewMetadataFacts: state.viewMetadata?.facts?.length || 0
  };

  // Debug: Show how many dimensions/facts have table info
  if (state.viewMetadata) {
    const dimsWithTable = state.viewMetadata.dimensions?.filter(d => d.logicalTable).length || 0;
    const factsWithTable = state.viewMetadata.facts?.filter(f => f.logicalTable).length || 0;
  }

  // Get logical tables of selected facts
  const factTables = new Set();
  facts.forEach(f => {
    const factMeta = state.viewMetadata.facts.find(fm => fm.name === f.name);
    if (factMeta && factMeta.logicalTable) {
      factTables.add(factMeta.logicalTable);
    } else if (factMeta) {
    }
  });

  // Get logical tables of selected dimensions
  const dimensionTables = new Set();
  dimensions.forEach(d => {
    const dimMeta = state.viewMetadata.dimensions.find(dm => dm.name === d.name);
    if (dimMeta && dimMeta.logicalTable) {
      dimensionTables.add(dimMeta.logicalTable);
    } else if (dimMeta) {
    }
  });

  // Log validation plan
  if (hasDimensions || hasMetrics || hasFacts) {
    const planMsg = `📋 Validation Plan: ${dimensions.length} dimensions, ${metrics.length} metrics, ${facts.length} facts selected`;
    if (hasFacts) {
    }
    if (hasDimensions) {
    }
    if (hasMetrics) {
    }
    if (hasDimensions && !hasMetrics) {
    }
  }

  let disabledCount = { dimension: 0, metric: 0, fact: 0 };

  for (const item of allItems) {
    const fieldName = item.dataset.fieldName;
    const fieldType = item.dataset.fieldType;

    let shouldDisable = false;
    let disabledReason = '';

    // Disable facts if metrics are selected
    if (fieldType === 'fact' && hasMetrics) {
      shouldDisable = true;
      disabledReason = 'Cannot use facts with metrics';
    }

    // Disable metrics if facts are selected
    if (fieldType === 'metric' && hasFacts) {
      shouldDisable = true;
      disabledReason = 'Cannot use metrics with facts';
    }

    // NEW: When facts are selected, block dimensions from other tables
    if (fieldType === 'dimension' && hasFacts && factTables.size > 0) {
      const checkMsg = `🔍 Checking dimension "${fieldName}" with facts from tables: ${Array.from(factTables).join(', ')}`;

      const dimMeta = state.viewMetadata.dimensions.find(dm => dm.name === fieldName);
      if (dimMeta && dimMeta.logicalTable) {
        if (!factTables.has(dimMeta.logicalTable)) {
          shouldDisable = true;
          const tableNames = Array.from(factTables).join(', ');
          disabledReason = `When using facts from '${tableNames}', dimensions must be from the same table. This dimension is from '${dimMeta.logicalTable}'.`;

          const disableMsg = `❌ DISABLED: dimension "${fieldName}" from table '${dimMeta.logicalTable}' (facts are from '${tableNames}')`;
        } else {
          const compatMsg = `✅ Compatible: dimension "${fieldName}" from table '${dimMeta.logicalTable}' matches fact table`;
        }
      }
    }

    // NEW: When facts are selected, block facts from other tables
    if (fieldType === 'fact' && hasFacts && factTables.size > 0) {
      const checkMsg = `🔍 Checking fact "${fieldName}" with existing facts from tables: ${Array.from(factTables).join(', ')}`;

      const factMeta = state.viewMetadata.facts.find(fm => fm.name === fieldName);
      if (factMeta && factMeta.logicalTable) {
        if (!factTables.has(factMeta.logicalTable)) {
          shouldDisable = true;
          const tableNames = Array.from(factTables).join(', ');
          disabledReason = `Facts must come from the same table. Selected facts are from '${tableNames}', but this fact is from '${factMeta.logicalTable}'.`;

          const disableMsg = `❌ DISABLED: fact "${fieldName}" from table '${factMeta.logicalTable}' (selected facts are from '${tableNames}')`;
        } else {
          const compatMsg = `✅ Compatible: fact "${fieldName}" from table '${factMeta.logicalTable}' matches selected facts`;
        }
      }
    }

    // Disable incompatible dimensions if metrics are selected
    // Constraint: Dimension's logical table must be related to metric's logical table
    // and have equal or lower granularity (validated via SHOW SEMANTIC DIMENSIONS FOR METRIC)
    if (fieldType === 'dimension' && hasMetrics && !shouldDisable) {
      const checkMsg = `🔍 Checking dimension "${fieldName}" with ${metrics.length} metrics`;

      // Check if dimension is compatible with all metrics
      let compatibleWithAll = true;
      let incompatibleMetric = null;
      for (const metric of metrics) {
        const allowed = await isDimensionAllowedForMetric(fieldName, metric.name);
        const resultMsg = `  → "${fieldName}" + "${metric.name}": ${allowed ? '✅ Compatible' : '❌ Incompatible'}`;
        if (!allowed) {
          compatibleWithAll = false;
          incompatibleMetric = metric.name;
          break;
        }
      }
      if (!compatibleWithAll) {
        shouldDisable = true;
        disabledReason = `Dimension "${fieldName}" cannot be used with metric "${incompatibleMetric}". The dimension's logical table must be related to the metric's table and have equal or lower granularity.`;
        const disableMsg = `❌ DISABLED: dimension "${fieldName}" (incompatible with "${incompatibleMetric}")`;
      } else {
        const compatMsg = `✅ Compatible: dimension "${fieldName}" works with all metrics`;
      }
    }

    // NEW: When dimensions are selected (and no facts), disable incompatible metrics
    // This ensures bidirectional validation: dimensions check metrics, and metrics check dimensions
    if (fieldType === 'metric' && hasDimensions && !hasFacts && !shouldDisable) {
      const checkMsg = `🔍 Checking metric "${fieldName}" with ${dimensions.length} dimensions`;

      // Check if this metric is compatible with all selected dimensions
      let compatibleWithAll = true;
      let incompatibleDimension = null;
      for (const dimension of dimensions) {
        const allowed = await isDimensionAllowedForMetric(dimension.name, fieldName);
        const resultMsg = `  → "${fieldName}" + "${dimension.name}": ${allowed ? '✅ Compatible' : '❌ Incompatible'}`;
        if (!allowed) {
          compatibleWithAll = false;
          incompatibleDimension = dimension.name;
          break;
        }
      }
      if (!compatibleWithAll) {
        shouldDisable = true;
        disabledReason = `Metric "${fieldName}" cannot be used with dimension "${incompatibleDimension}". The metric's logical table must be related to the dimension's table.`;
        const disableMsg = `❌ DISABLED: metric "${fieldName}" (incompatible with "${incompatibleDimension}")`;
      } else {
        const compatMsg = `✅ Compatible: metric "${fieldName}" works with all dimensions`;
      }
    }

    // NEW: When dimensions are selected (and no metrics/facts), disable facts based on table compatibility
    // Rule: Facts require dimensions from the SAME table. If dimensions are from multiple tables, disable all facts.
    if (fieldType === 'fact' && hasDimensions && !hasMetrics && !hasFacts && !shouldDisable) {
      const checkMsg = `🔍 Checking fact "${fieldName}" with dimensions from tables: ${Array.from(dimensionTables).join(', ')}`;

      const factMeta = state.viewMetadata.facts.find(fm => fm.name === fieldName);
      if (factMeta && factMeta.logicalTable) {
        // If dimensions are from multiple tables, disable all facts
        if (dimensionTables.size > 1) {
          shouldDisable = true;
          const dimTableNames = Array.from(dimensionTables).join(', ');
          disabledReason = `Cannot use facts when dimensions are from multiple tables (${dimTableNames}). Facts require all dimensions to be from the same table. Consider using Metrics instead.`;

          const disableMsg = `❌ DISABLED: fact "${fieldName}" (dimensions are from multiple tables: ${dimTableNames})`;
        }
        // If dimensions are from one table, only allow facts from that table
        else if (dimensionTables.size === 1) {
          if (!dimensionTables.has(factMeta.logicalTable)) {
            shouldDisable = true;
            const dimTableName = Array.from(dimensionTables)[0];
            disabledReason = `Fact "${fieldName}" is from table '${factMeta.logicalTable}', but selected dimensions are from table '${dimTableName}'. Facts and dimensions must be from the same table.`;

            const disableMsg = `❌ DISABLED: fact "${fieldName}" from table '${factMeta.logicalTable}' (dimensions are from '${dimTableName}')`;
          } else {
            const compatMsg = `✅ Compatible: fact "${fieldName}" from table '${factMeta.logicalTable}' matches dimension table`;
          }
        }
      }
    }

    // Update UI
    const label = item.querySelector('.field-label');
    if (shouldDisable) {
      item.classList.add('disabled');
      if (label) label.setAttribute('draggable', 'false');
      item.title = disabledReason;
      disabledCount[fieldType]++;
    } else {
      item.classList.remove('disabled');
      if (label) label.setAttribute('draggable', 'true');
      item.title = '';
    }
  }

  const summary = {
    dimensions: disabledCount.dimension,
    metrics: disabledCount.metric,
    facts: disabledCount.fact,
    total: disabledCount.dimension + disabledCount.metric + disabledCount.fact
  };
}

// Expose functions and objects globally for inline event handlers
window.SnowflakeAuth = SnowflakeAuth;
window.PKCEHelper = PKCEHelper;
window.SnowflakeAPI = SnowflakeAPI;
window.AdhocMetricsManager = AdhocMetricsManager;
window.appState = state;
window.selectAllFields = selectAllFields;
window.deselectAllFields = deselectAllFields;
window.removeChip = removeChip;
window.removeCustomFilter = removeCustomFilter;
window.getCustomFilters = getCustomFilters;
