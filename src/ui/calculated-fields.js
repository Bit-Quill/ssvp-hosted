/**
 * Adhoc Metrics (Calculated Fields) Component
 * Allows users to create and update custom metrics based on existing fields
 */

const AdhocMetricsManager = {
  adhocMetrics: [],
  currentEditingMetricId: null,
  availableFields: {
    dimensions: [],
    metrics: [],
    facts: []
  },

  /**
   * Initialize the adhoc metrics manager
   */
  init(fields) {
    this.availableFields = fields;
    this.adhocMetrics = [];
    this.renderAdhocMetricsList();
  },

  /**
   * Update available fields when semantic view changes
   */
  updateAvailableFields(fields) {
    this.availableFields = fields;
  },

  /**
   * Show dialog to create a new adhoc metric
   */
  showCreateDialog() {
    this.currentEditingMetricId = null;
    this.showDialog({
      title: 'Create Adhoc Metric',
      name: '',
      expression: '',
      description: '',
      buttonText: 'Create'
    });
  },

  /**
   * Show dialog to update an existing adhoc metric
   */
  showUpdateDialog(metricId) {
    const metric = this.adhocMetrics.find(m => m.id === metricId);
    if (!metric) {
      throw new Error('Adhoc metric not found');
    }

    this.currentEditingMetricId = metricId;
    this.showDialog({
      title: 'Update Adhoc Metric',
      name: metric.name,
      expression: metric.expression,
      description: metric.description || '',
      buttonText: 'Update'
    });
  },

  /**
   * Show the adhoc metric dialog (for both Create and Update)
   */
  showDialog(config) {
    const { title, name, expression, description, buttonText } = config;

    // Create dialog overlay
    const overlay = document.createElement('div');
    overlay.className = 'adhoc-metric-overlay';
    overlay.id = 'adhoc-metric-overlay';

    // Create dialog content
    overlay.innerHTML = `
      <div class="adhoc-metric-dialog">
        <div class="dialog-header">
          <h3>${title}</h3>
          <button class="close-btn" id="adhoc-close-btn">×</button>
        </div>

        <div class="dialog-body">
          <div class="form-group">
            <label for="adhoc-name">
              Metric Name <span class="required">*</span>
            </label>
            <input
              type="text"
              id="adhoc-name"
              class="input-field"
              placeholder="e.g., Total_Revenue"
              value="${name}"
              required
            />
            <small class="help-text">
              Use letters, numbers, and underscores only
            </small>
          </div>

          <div class="form-group">
            <label for="adhoc-expression">
              Expression <span class="required">*</span>
            </label>
            <textarea
              id="adhoc-expression"
              class="input-field"
              rows="4"
              placeholder="e.g., SUM(sales_amount) * 1.1"
              required
            >${expression}</textarea>
            <small class="help-text">
              Use SQL aggregate functions (SUM, AVG, COUNT, etc.) and existing fields
            </small>
          </div>

          <div class="form-group">
            <label for="adhoc-description">Description (Optional)</label>
            <input
              type="text"
              id="adhoc-description"
              class="input-field"
              placeholder="Describe what this metric calculates"
              value="${description}"
            />
          </div>

          <div class="available-fields-section">
            <details>
              <summary>=Ë Available Fields (Click to insert)</summary>
              <div class="fields-grid">
                <div class="field-column">
                  <strong>Dimensions</strong>
                  <ul id="adhoc-dimensions-list"></ul>
                </div>
                <div class="field-column">
                  <strong>Metrics</strong>
                  <ul id="adhoc-metrics-list"></ul>
                </div>
                <div class="field-column">
                  <strong>Facts</strong>
                  <ul id="adhoc-facts-list"></ul>
                </div>
              </div>
            </details>
          </div>

          <div id="validation-message" class="validation-message" style="display: none;"></div>
        </div>

        <div class="dialog-footer">
          <button id="adhoc-cancel-btn" class="secondary-button">Cancel</button>
          <button id="adhoc-save-btn" class="primary-button">${buttonText}</button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);

    // Populate available fields
    this.populateAvailableFieldsInDialog();

    // Bind events
    document.getElementById('adhoc-close-btn').addEventListener('click', () => this.closeDialog());
    document.getElementById('adhoc-cancel-btn').addEventListener('click', () => this.closeDialog());
    document.getElementById('adhoc-save-btn').addEventListener('click', () => this.handleSave());

    // Real-time validation
    const expressionInput = document.getElementById('adhoc-expression');
    expressionInput.addEventListener('input', () => this.validateExpression());

    const nameInput = document.getElementById('adhoc-name');
    nameInput.addEventListener('input', () => this.validateName());
  },

  /**
   * Populate available fields in dialog
   */
  populateAvailableFieldsInDialog() {
    const dimensionsList = document.getElementById('adhoc-dimensions-list');
    const metricsList = document.getElementById('adhoc-metrics-list');
    const factsList = document.getElementById('adhoc-facts-list');

    // Clear lists
    dimensionsList.innerHTML = '';
    metricsList.innerHTML = '';
    factsList.innerHTML = '';

    // Populate dimensions
    this.availableFields.dimensions.forEach(dim => {
      const li = document.createElement('li');
      li.textContent = dim.name;
      li.className = 'insertable-field';
      li.title = `Click to insert: ${dim.name}`;
      li.addEventListener('click', () => this.insertFieldIntoExpression(dim.name));
      dimensionsList.appendChild(li);
    });

    // Populate metrics
    this.availableFields.metrics.forEach(metric => {
      const li = document.createElement('li');
      li.textContent = metric.name;
      li.className = 'insertable-field';
      li.title = `Click to insert: ${metric.name}`;
      li.addEventListener('click', () => this.insertFieldIntoExpression(metric.name));
      metricsList.appendChild(li);
    });

    // Populate facts
    this.availableFields.facts.forEach(fact => {
      const li = document.createElement('li');
      li.textContent = fact.name;
      li.className = 'insertable-field';
      li.title = `Click to insert: ${fact.name}`;
      li.addEventListener('click', () => this.insertFieldIntoExpression(fact.name));
      factsList.appendChild(li);
    });
  },

  /**
   * Insert field name into expression textarea
   */
  insertFieldIntoExpression(fieldName) {
    const textarea = document.getElementById('adhoc-expression');
    const cursorPos = textarea.selectionStart;
    const textBefore = textarea.value.substring(0, cursorPos);
    const textAfter = textarea.value.substring(cursorPos);

    textarea.value = textBefore + fieldName + textAfter;
    textarea.focus();
    textarea.selectionStart = textarea.selectionEnd = cursorPos + fieldName.length;

    this.validateExpression();
  },

  /**
   * Validate metric name
   */
  validateName() {
    const nameInput = document.getElementById('adhoc-name');
    const name = nameInput.value.trim();

    // Name validation: letters, numbers, underscores only
    const validNameRegex = /^[A-Za-z_][A-Za-z0-9_]*$/;

    if (!name) {
      this.showValidationMessage('Metric name is required', 'error');
      return false;
    }

    if (!validNameRegex.test(name)) {
      this.showValidationMessage(
        'Metric name must start with a letter or underscore, and contain only letters, numbers, and underscores',
        'error'
      );
      return false;
    }

    // Check for duplicates (exclude current editing metric)
    const isDuplicate = this.adhocMetrics.some(
      m => m.name === name && m.id !== this.currentEditingMetricId
    );

    if (isDuplicate) {
      this.showValidationMessage('A metric with this name already exists', 'error');
      return false;
    }

    this.hideValidationMessage();
    return true;
  },

  /**
   * Validate expression
   */
  validateExpression() {
    const expressionInput = document.getElementById('adhoc-expression');
    const expression = expressionInput.value.trim();

    if (!expression) {
      this.showValidationMessage('Expression is required', 'error');
      return false;
    }

    // Check for SQL injection patterns (basic validation)
    const dangerousPatterns = [';', '--', '/*', '*/', 'DROP', 'DELETE', 'INSERT', 'UPDATE', 'CREATE', 'ALTER'];
    const upperExpression = expression.toUpperCase();

    for (const pattern of dangerousPatterns) {
      if (upperExpression.includes(pattern.toUpperCase())) {
        this.showValidationMessage(
          `Expression contains potentially dangerous pattern: "${pattern}"`,
          'error'
        );
        return false;
      }
    }

    // Validate that expression uses valid fields
    const tokens = expression.match(/\b\w+\b/g) || [];
    const allFields = [
      ...this.availableFields.dimensions.map(d => d.name),
      ...this.availableFields.metrics.map(m => m.name),
      ...this.availableFields.facts.map(f => f.name)
    ];

    // SQL keywords and functions to ignore
    const sqlKeywords = [
      'SUM', 'AVG', 'COUNT', 'MIN', 'MAX', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END',
      'AND', 'OR', 'NOT', 'AS', 'FROM', 'WHERE', 'GROUP', 'BY', 'ORDER', 'HAVING',
      'DISTINCT', 'CAST', 'COALESCE', 'NULLIF', 'IFF', 'IFNULL'
    ];

    for (const token of tokens) {
      const upperToken = token.toUpperCase();

      // Skip SQL keywords and functions
      if (sqlKeywords.includes(upperToken)) {
        continue;
      }

      // Skip numbers
      if (!isNaN(token)) {
        continue;
      }

      // Must be a valid field name
      if (!allFields.includes(token)) {
        this.showValidationMessage(
          `Unknown field: "${token}". Use fields from Available Fields below.`,
          'warning'
        );
        return false;
      }
    }

    this.showValidationMessage('Expression looks good!', 'success');
    return true;
  },

  /**
   * Show validation message
   */
  showValidationMessage(message, type = 'error') {
    const validationDiv = document.getElementById('validation-message');
    validationDiv.textContent = message;
    validationDiv.className = `validation-message ${type}`;
    validationDiv.style.display = 'block';
  },

  /**
   * Hide validation message
   */
  hideValidationMessage() {
    const validationDiv = document.getElementById('validation-message');
    validationDiv.style.display = 'none';
  },

  /**
   * Handle Save button click (Create or Update)
   */
  handleSave() {
    const name = document.getElementById('adhoc-name').value.trim();
    const expression = document.getElementById('adhoc-expression').value.trim();
    const description = document.getElementById('adhoc-description').value.trim();

    // Validate
    const nameValid = this.validateName();
    const expressionValid = this.validateExpression();

    if (!nameValid || !expressionValid) {
      return;
    }

    if (this.currentEditingMetricId) {
      // Update existing metric
      this.updateAdhocMetric(this.currentEditingMetricId, { name, expression, description });
    } else {
      // Create new metric
      this.createAdhocMetric({ name, expression, description });
    }

    this.closeDialog();
  },

  /**
   * Create a new adhoc metric
   */
  createAdhocMetric({ name, expression, description }) {
    const newMetric = {
      id: `adhoc_${Date.now()}`,
      name,
      expression,
      description: description || '',
      type: 'adhoc_metric',
      createdAt: new Date().toISOString()
    };

    this.adhocMetrics.push(newMetric);
    this.renderAdhocMetricsList();

    // Trigger event for parent component
    this.triggerEvent('adhoc-metric-created', newMetric);

    // Show success message
    this.showSuccessToast(`Adhoc metric "${name}" created successfully`);
  },

  /**
   * Update an existing adhoc metric
   */
  updateAdhocMetric(metricId, { name, expression, description }) {
    const metric = this.adhocMetrics.find(m => m.id === metricId);
    if (!metric) {
      throw new Error('Adhoc metric not found');
    }

    metric.name = name;
    metric.expression = expression;
    metric.description = description;
    metric.updatedAt = new Date().toISOString();

    this.renderAdhocMetricsList();

    // Trigger event for parent component
    this.triggerEvent('adhoc-metric-updated', metric);

    // Show success message
    this.showSuccessToast(`Adhoc metric "${name}" updated successfully`);
  },

  /**
   * Delete an adhoc metric
   */
  deleteAdhocMetric(metricId) {
    const metric = this.adhocMetrics.find(m => m.id === metricId);
    if (!metric) {
      return;
    }

    if (!confirm(`Are you sure you want to delete adhoc metric "${metric.name}"?`)) {
      return;
    }

    this.adhocMetrics = this.adhocMetrics.filter(m => m.id !== metricId);
    this.renderAdhocMetricsList();

    // Trigger event for parent component
    this.triggerEvent('adhoc-metric-deleted', metric);

    this.showSuccessToast(`Adhoc metric "${metric.name}" deleted`);
  },

  /**
   * Render the adhoc metrics list
   */
  renderAdhocMetricsList() {
    const container = document.getElementById('adhoc-metrics-container');
    if (!container) return;

    if (this.adhocMetrics.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <p>No adhoc metrics created yet.</p>
          <small>Click "Create Adhoc Metric" to get started.</small>
        </div>
      `;
      return;
    }

    container.innerHTML = `
      <ul class="adhoc-metrics-list">
        ${this.adhocMetrics.map(metric => `
          <li class="adhoc-metric-item" data-metric-id="${metric.id}">
            <div class="metric-info">
              <div class="metric-name">
                <span class="metric-icon">’</span>
                <strong>${metric.name}</strong>
              </div>
              <div class="metric-expression">
                <code>${metric.expression}</code>
              </div>
              ${metric.description ? `
                <div class="metric-description">
                  <small>${metric.description}</small>
                </div>
              ` : ''}
            </div>
            <div class="metric-actions">
              <button class="icon-button" data-action="edit" data-metric-id="${metric.id}" title="Edit">
                
              </button>
              <button class="icon-button" data-action="delete" data-metric-id="${metric.id}" title="Delete">
                =Ñ
              </button>
            </div>
          </li>
        `).join('')}
      </ul>
    `;

    // Bind action buttons
    container.querySelectorAll('button[data-action="edit"]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.showUpdateDialog(btn.dataset.metricId);
      });
    });

    container.querySelectorAll('button[data-action="delete"]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.deleteAdhocMetric(btn.dataset.metricId);
      });
    });
  },

  /**
   * Close the dialog
   */
  closeDialog() {
    const overlay = document.getElementById('adhoc-metric-overlay');
    if (overlay) {
      overlay.remove();
    }
    this.currentEditingMetricId = null;
  },

  /**
   * Get all adhoc metrics
   */
  getAdhocMetrics() {
    return [...this.adhocMetrics];
  },

  /**
   * Clear all adhoc metrics
   */
  clearAdhocMetrics() {
    this.adhocMetrics = [];
    this.renderAdhocMetricsList();
  },

  /**
   * Trigger custom event
   */
  triggerEvent(eventName, data) {
    const event = new CustomEvent(eventName, { detail: data });
    document.dispatchEvent(event);
  },

  /**
   * Show success toast message
   */
  showSuccessToast(message) {
    const toast = document.createElement('div');
    toast.className = 'success-toast';
    toast.textContent = message;
    document.body.appendChild(toast);

    setTimeout(() => {
      toast.classList.add('show');
    }, 10);

    setTimeout(() => {
      toast.classList.remove('show');
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }
};

export default AdhocMetricsManager;
