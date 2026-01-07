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
  viewMetadata: null
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

  const btnCloseView = document.getElementById('btn-close-view');
  if (btnCloseView) {
    btnCloseView.addEventListener('click', handleCloseView);
  }

  const btnLoadToExcel = document.getElementById('btn-load-to-excel');
  if (btnLoadToExcel) {
    btnLoadToExcel.addEventListener('click', handleLoadToExcel);
  }

  const btnPreviewQuery = document.getElementById('btn-preview-query');
  if (btnPreviewQuery) {
    btnPreviewQuery.addEventListener('click', handlePreviewQuery);
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

  const clientIdInput = document.getElementById('client-id-input');
  const roleInput = document.getElementById('role-input');

  const clientId = clientIdInput ? clientIdInput.value.trim() || null : null;
  const role = roleInput ? roleInput.value.trim() || null : null;

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

    await loadWarehousesAndRoles();
    showRoleSelection();

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

    const rolesResult = await SnowflakeAuth.executeStatement('SHOW ROLES');
    state.availableRoles = rolesResult.data.map(row => row[1]);

    populateWizardRoleDropdown();
    hideLoading();

  } catch (error) {
    hideLoading();
    throw new Error('Failed to load roles. You may not have permission.');
  }
}

function populateWizardRoleDropdown() {
  const select = document.getElementById('role-select');
  select.innerHTML = '';

  state.availableRoles.forEach(role => {
    const option = document.createElement('option');
    option.value = role;
    option.textContent = role;
    select.appendChild(option);
  });
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
    const warehousesResult = await SnowflakeAuth.executeStatement('SHOW WAREHOUSES', { role: role });
    state.availableWarehouses = warehousesResult.data.map(row => row[0]);
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
    const result = await SnowflakeAuth.executeStatement(
      'SELECT CURRENT_WAREHOUSE() AS WH, CURRENT_ROLE() AS ROLE'
    );

    const verifiedWarehouse = result.data[0][0];
    const verifiedRole = result.data[0][1];

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

    const result = await SnowflakeAuth.executeStatement(
      'SELECT CURRENT_WAREHOUSE() AS WH, CURRENT_ROLE() AS ROLE'
    );

    const verifiedWarehouse = result.data[0][0];
    const verifiedRole = result.data[0][1];

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
  await loadViewMetadata(fullyQualifiedName);
}

async function loadViewMetadata(fullyQualifiedName) {
  try {
    showLoading('Loading view metadata...');
    const metadata = await SnowflakeAPI.describeSemanticView(fullyQualifiedName);
    state.viewMetadata = metadata;

    document.getElementById('selected-view-name').textContent = fullyQualifiedName;

    displayFieldList('dimensions-list', metadata.dimensions, 'dimension');
    document.getElementById('dimensions-count').textContent = metadata.dimensions.length;

    displayFieldList('metrics-list', metadata.metrics, 'metric');
    document.getElementById('metrics-count').textContent = metadata.metrics.length;

    displayFieldList('facts-list', metadata.facts, 'fact');
    document.getElementById('facts-count').textContent = metadata.facts.length;

    if (metadata.nonEnforcedFilters && metadata.nonEnforcedFilters.length > 0) {
      displayRecommendedFilters(metadata.nonEnforcedFilters);
      document.getElementById('filters-section').classList.remove('hidden');
      document.getElementById('filters-count').textContent = metadata.nonEnforcedFilters.length;
    } else {
      document.getElementById('filters-section').classList.add('hidden');
    }

    AdhocMetricsManager.init({
      dimensions: metadata.dimensions,
      metrics: metadata.metrics,
      facts: metadata.facts
    });

    document.getElementById('view-metadata-section').classList.remove('hidden');

    hideLoading();
  } catch (error) {
    hideLoading();
    showError(`Failed to load view metadata: ${error.message}`);
  }
}

function displayFieldList(containerId, fields, fieldType) {
  const container = document.getElementById(containerId);

  if (fields.length === 0) {
    container.innerHTML = '<li class="empty-item">No fields available</li>';
    return;
  }

  container.innerHTML = fields.map(field => `
    <li class="field-item" data-field-name="${field.name}" data-field-type="${fieldType}">
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
    </li>
  `).join('');
}

function displayRecommendedFilters(filters) {
  const container = document.getElementById('recommended-filters-list');

  container.innerHTML = filters.map(filter => `
    <li class="filter-item">
      <div class="filter-info">
        <strong>${filter.name}</strong>
        ${filter.description ? `<p>${filter.description}</p>` : ''}
        <code>${filter.expression}</code>
      </div>
      <small class="filter-hint">💡 Recommended for performance</small>
    </li>
  `).join('');
}

function handleCloseView() {
  state.selectedView = null;
  state.viewMetadata = null;
  document.getElementById('view-metadata-section').classList.add('hidden');
  AdhocMetricsManager.clearAdhocMetrics();
}

async function handleLoadToExcel() {
  if (!state.viewMetadata) {
    showError('Please select a semantic view first');
    return;
  }

  try {
    showLoading('Loading data to Excel...');

    // For MVP, load all dimensions and metrics (basic implementation)
    const dimensions = state.viewMetadata.dimensions.map(d => d.name);
    const metrics = state.viewMetadata.metrics.map(m => m.name);
    const adhocMetrics = AdhocMetricsManager.getAdhocMetrics();

    if (dimensions.length === 0 && metrics.length === 0 && adhocMetrics.length === 0) {
      hideLoading();
      showError('Please select at least one field to load');
      return;
    }

    const result = await SnowflakeAPI.executeSemanticQuery({
      semanticView: state.selectedView,
      dimensions: dimensions.slice(0, 3), // Limit for MVP
      metrics: metrics.slice(0, 3), // Limit for MVP
      adhocMetrics: adhocMetrics,
      limit: 1000
    });

    // Write to Excel
    await writeToExcel(result);

    hideLoading();
    showSuccess('Data loaded to Excel successfully!');
  } catch (error) {
    hideLoading();
    showError(`Failed to load data: ${error.message}`);
  }
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

async function handlePreviewQuery() {
  if (!state.viewMetadata) {
    showError('Please select a semantic view first');
    return;
  }

  const dimensions = state.viewMetadata.dimensions.map(d => d.name);
  const metrics = state.viewMetadata.metrics.map(m => m.name);
  const adhocMetrics = AdhocMetricsManager.getAdhocMetrics();

  let query = `SELECT ${[...dimensions.slice(0, 3), ...metrics.slice(0, 3)].join(', ')}\nFROM ${state.selectedView}`;

  if (adhocMetrics.length > 0) {
    const adhocFields = adhocMetrics.map(m => `${m.expression} AS ${m.name}`).join(',\n  ');
    query = query.replace('FROM', `,\n  ${adhocFields}\nFROM`);
  }

  if (dimensions.length > 0 && (metrics.length > 0 || adhocMetrics.length > 0)) {
    query += `\nGROUP BY ${dimensions.slice(0, 3).join(', ')}`;
  }

  query += '\nLIMIT 1000';

  alert(`Preview Query:\n\n${query}`);
}

window.SnowflakeAuth = SnowflakeAuth;
window.PKCEHelper = PKCEHelper;
window.SnowflakeAPI = SnowflakeAPI;
window.AdhocMetricsManager = AdhocMetricsManager;
window.appState = state;
