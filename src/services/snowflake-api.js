/**
 * Snowflake API Service
 * Provides methods for interacting with Snowflake Semantic Views
 */

import SnowflakeAuth from './snowflake-auth.js';

const QUERY_CONFIG = {
  DEFAULT_TIMEOUT_SECONDS: 300,
  DEFAULT_RESULT_LIMIT: 10000,
  DEFAULT_DISTINCT_VALUES_LIMIT: 1000
};

const POLLING_CONFIG = {
  MAX_ATTEMPTS: 60,
  INITIAL_DELAY_MS: 500,
  MAX_DELAY_MS: 5000,
  BACKOFF_MULTIPLIER: 1.5
};

const SnowflakeAPI = {
  /**
   * Execute a SQL query via Snowflake SQL API
   */
  async executeQuery(sql, options = {}) {
    try {
      const token = await SnowflakeAuth.getAccessToken();
      const { timeout = QUERY_CONFIG.DEFAULT_TIMEOUT_SECONDS, warehouse = null, role = null } = options;

      // Make direct calls to Snowflake SQL API v2
      const requestBody = {
        statement: sql,
        timeout: timeout,
        resultSetMetaData: { format: 'json' },
        parameters: {
          query_tag: 'semantic_view_excel_plugin'
        }
      };

      if (warehouse || SnowflakeAuth.config.warehouse) {
        requestBody.warehouse = warehouse || SnowflakeAuth.config.warehouse;
      }

      if (role || SnowflakeAuth.config.role) {
        requestBody.role = role || SnowflakeAuth.config.role;
      }

      const headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': `Bearer ${token}`
      };

      const response = await fetch(
        `https://${SnowflakeAuth.config.account}.snowflakecomputing.com/api/v2/statements`,
        {
          method: 'POST',
          headers: headers,
          body: JSON.stringify(requestBody)
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        let errorMessage = 'Query execution failed';
        try {
          const error = JSON.parse(errorText);
          errorMessage = error.message || error.error || errorText;
        } catch (e) {
          errorMessage = errorText;
        }
        throw new Error(errorMessage);
      }

      const result = await response.json();

      // Check for immediate success
      if (result.resultSetMetaData && result.data) {
        return this.parseResults(result);
      }

      // Handle async execution (query still running)
      if (result.statementStatusUrl || result.statementHandle) {
        let statusUrl;

        if (result.statementStatusUrl) {
          // Snowflake returned a status URL - check if it's relative or absolute
          if (result.statementStatusUrl.startsWith('http')) {
            // Already absolute URL
            statusUrl = result.statementStatusUrl;
          } else {
            // Relative URL - prepend hostname
            statusUrl = `https://${SnowflakeAuth.config.account}.snowflakecomputing.com${result.statementStatusUrl}`;
          }
        } else {
          // No status URL, construct from statement handle
          statusUrl = `https://${SnowflakeAuth.config.account}.snowflakecomputing.com/api/v2/statements/${result.statementHandle}`;
        }

        return await this.pollForResults(statusUrl);
      }

      // If we get here, something unexpected happened
      console.error('Unexpected query result format:', result);
      throw new Error('Unexpected query result format');
    } catch (error) {
      throw new Error(`Failed to execute query: ${error.message}`);
    }
  },

  /**
   * Poll for async query results
   */
  async pollForResults(statusUrl) {
    let attempts = 0;

    while (attempts < POLLING_CONFIG.MAX_ATTEMPTS) {
      try {
        const token = await SnowflakeAuth.getAccessToken();

        // SQL API v2 uses Bearer format for OAuth and PAT
        const headers = {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Authorization': `Bearer ${token}`,
          'X-Snowflake-Authorization-Token-Type': 'OAUTH'
        };

        // Add partition parameter to request full result set
        const pollUrl = statusUrl.includes('?') ? statusUrl : `${statusUrl}?partition=0`;

        const response = await fetch(pollUrl, {
          method: 'GET',
          headers: headers
        });

        if (!response.ok) {
          const errorText = await response.text();
          console.error('Poll failed:', response.status, errorText);

          // 404 errors - provide detailed diagnostic info
          if (response.status === 404) {
            throw new Error(
              `Statement not found (404)\n\n` +
              `Base URL: ${statusUrl}\n` +
              `Poll URL: ${pollUrl}\n` +
              `Account: ${SnowflakeAuth.config.account}\n` +
              `Attempt: ${attempts + 1}/${POLLING_CONFIG.MAX_ATTEMPTS}\n\n` +
              `The statement may have expired or the URL is incorrect.\n` +
              `Response: ${errorText.substring(0, 200)}`
            );
          }

          throw new Error(`Failed to poll query status: ${response.status}\n${errorText.substring(0, 300)}`);
        }

        const status = await response.json();

        // Check various success status formats
        if (status.resultSetMetaData && status.data) {
          // Query completed with results
          return this.parseResults(status);
        }

        // Check execution status
        const execStatus = status.statementStatusUrl ? 'running' : status.message;

        if (execStatus === 'error' || status.sqlState === 'ERROR') {
          throw new Error(status.message || 'Query execution failed');
        }

        // Calculate exponential backoff delay
        const delay = Math.min(
          POLLING_CONFIG.INITIAL_DELAY_MS * Math.pow(POLLING_CONFIG.BACKOFF_MULTIPLIER, attempts),
          POLLING_CONFIG.MAX_DELAY_MS
        );
        await new Promise(resolve => setTimeout(resolve, delay));
        attempts++;
      } catch (error) {
        throw error;
      }
    }

    throw new Error(`Query timed out after ${POLLING_CONFIG.MAX_ATTEMPTS} polling attempts`);
  },

  /**
   * Parse query results into array of objects
   */
  parseResults(result) {
    try {
      if (!result.data || result.data.length === 0) {
        return { columns: [], rows: [] };
      }

      if (!result.resultSetMetaData || !result.resultSetMetaData.rowType) {
        console.error('Invalid result format:', result);
        throw new Error('Invalid result format: missing resultSetMetaData');
      }

      const columns = result.resultSetMetaData.rowType.map(col => ({
        name: col.name,
        type: col.type,
        nullable: col.nullable
      }));

      const rows = result.data.map(row => {
        const obj = {};
        columns.forEach((col, index) => {
          obj[col.name] = row[index];
        });
        return obj;
      });

      return { columns, rows };
    } catch (error) {
      console.error('Parse results error:', error, 'Result:', result);
      throw new Error(`Failed to parse results: ${error.message}`);
    }
  },

  /**
   * Get list of databases
   */
  async getDatabases() {
    try {
      const result = await this.executeQuery('SHOW DATABASES');
      return result.rows.map(row => ({
        name: row.name,
        owner: row.owner,
        created_on: row.created_on
      }));
    } catch (error) {
      throw new Error(`Failed to get databases: ${error.message}`);
    }
  },

  /**
   * Get schemas in a database
   */
  async getSchemas(database) {
    try {
      const result = await this.executeQuery(`SHOW SCHEMAS IN DATABASE ${database}`);
      return result.rows.map(row => ({
        name: row.name,
        database: row.database_name,
        owner: row.owner
      }));
    } catch (error) {
      throw new Error(`Failed to get schemas: ${error.message}`);
    }
  },

  /**
   * Get semantic views in a schema
   * Implements: Catalog Browser requirement
   */
  async getSemanticViews(database, schema) {
    try {
      const sql = `SHOW SEMANTIC VIEWS IN SCHEMA ${database}.${schema}`;
      const result = await this.executeQuery(sql);

      return result.rows.map(row => ({
        name: row.name,
        database: row.database_name,
        schema: row.schema_name,
        owner: row.owner,
        created_on: row.created_on,
        comment: row.comment || ''
      }));
    } catch (error) {
      // If error is "object does not exist", return empty array
      if (error.message.includes('does not exist') || error.message.includes('Insufficient privileges')) {
        return [];
      }
      throw new Error(`Failed to get semantic views: ${error.message}`);
    }
  },

  /**
   * Describe a semantic view to get metadata
   * Returns: { dimensions, metrics, facts, nonEnforcedFilters }
   *
   * DESC SEMANTIC VIEW returns rows with:
   * - object_kind: COLUMN | MEASURE | DIMENSION | TIME_DIMENSION | METRIC | FACT (and others like TABLE, VIEW that we filter out)
   * - object_name: name of the object
   * - property: property name (SEMANTIC_TYPE, DATA_TYPE, DESCRIPTION, EXPRESSION, BASE_TABLE_NAME, etc.)
   * - property_value: the value
   */
  async describeSemanticView(fullyQualifiedName) {
    try {
      const sql = `DESC SEMANTIC VIEW ${fullyQualifiedName}`;
      const result = await this.executeQuery(sql);

      // First, let's see what object kinds exist
      const objectKinds = new Set();
      result.rows.forEach(row => {
        const objectKind = row.object_kind || row.OBJECT_KIND;
        objectKinds.add(objectKind);
      });

      // Group rows by object_name to reconstruct each field
      const objectsMap = new Map();
      const allObjectsMap = new Map(); // Track ALL objects for debugging

      result.rows.forEach(row => {
        const objectKind = row.object_kind || row.OBJECT_KIND;
        const objectName = row.object_name || row.OBJECT_NAME;
        const property = row.property || row.PROPERTY;
        const propertyValue = row.property_value || row.PROPERTY_VALUE;

        // Track all objects in allObjectsMap for debugging
        if (!allObjectsMap.has(objectName)) {
          allObjectsMap.set(objectName, {
            name: objectName,
            kind: objectKind,
            properties: {}
          });
        }
        allObjectsMap.get(objectName).properties[property] = propertyValue;

        // Only process field-related objects (not TABLE, VIEW, RELATIONSHIP, etc.)
        if (!['COLUMN', 'MEASURE', 'DIMENSION', 'TIME_DIMENSION', 'FILTER', 'METRIC', 'FACT'].includes(objectKind)) {
          return;
        }

        if (!objectsMap.has(objectName)) {
          objectsMap.set(objectName, {
            name: objectName,
            kind: objectKind,
            properties: {}
          });
        }

        objectsMap.get(objectName).properties[property] = propertyValue;
      });

      const dimensions = [];
      const metrics = [];
      const facts = [];
      const nonEnforcedFilters = [];

      // Convert grouped objects into dimensions/metrics/facts
      objectsMap.forEach((obj, name) => {
        const props = obj.properties;
        const semanticType = props.SEMANTIC_TYPE || props.semantic_type || props.COLUMN_TYPE;
        const dataType = props.DATA_TYPE || props.data_type || 'VARCHAR';
        const description = props.DESCRIPTION || props.description || '';
        const expression = props.EXPRESSION || props.expression || props.EXPR;
        const baseTable = props.BASE_TABLE_NAME || props.BASE_TABLE || props.TABLE_NAME || null;

        const field = {
          name: name,
          dataType: dataType,
          description: description
        };

        let categorizedAs = 'unknown';

        // Determine field type based on object_kind and semantic_type
        // Priority: DIMENSION/TIME_DIMENSION > MEASURE/METRIC > COLUMN/FACT

        // 1. Check for dimensions (explicit dimension objects)
        if (obj.kind === 'DIMENSION' || obj.kind === 'TIME_DIMENSION') {
          dimensions.push({
            ...field,
            logicalTable: baseTable
          });
          categorizedAs = 'dimension';
        }
        // 2. Check for metrics/measures (explicit measure objects or have expressions)
        else if (obj.kind === 'MEASURE' || obj.kind === 'METRIC') {
          metrics.push({
            ...field,
            expression: expression || null,
            logicalTable: baseTable
          });
          categorizedAs = 'metric';
        }
        // 3. Check for explicit FACT objects
        else if (obj.kind === 'FACT') {
          const isNumeric = this.isNumericType(dataType);
          facts.push({
            ...field,
            isNumeric: isNumeric,
            canBeMetric: isNumeric,
            logicalTable: baseTable
          });
          categorizedAs = 'fact';
        }
        // NOTE: The 'FILTER' case was removed because Snowflake semantic views
        // don't have a separate kind named 'FILTER'. Filters are identified via
        // labels on facts or dimensions, not as a separate kind.
        // TODO: Implement filter discovery based on labels when the feature is needed.
        // 4. COLUMN objects - check semantic type to determine if dimension or fact
        else if (obj.kind === 'COLUMN') {
          // If column has semantic type DIMENSION, it's a dimension
          if (semanticType === 'DIMENSION' || semanticType === 'TIME_DIMENSION') {
            dimensions.push({
              ...field,
              logicalTable: baseTable
            });
            categorizedAs = 'dimension (from COLUMN)';
          }
          // If column has semantic type MEASURE/METRIC, it's a metric
          else if (semanticType === 'MEASURE' || semanticType === 'METRIC') {
            metrics.push({
              ...field,
              expression: expression || null,
              logicalTable: baseTable
            });
            categorizedAs = 'metric (from COLUMN)';
          }
          // Otherwise, it's a fact (regular data column)
          else {
            const isNumeric = this.isNumericType(dataType);
            facts.push({
              ...field,
              isNumeric: isNumeric,
              canBeMetric: isNumeric,
              logicalTable: baseTable
            });
            categorizedAs = 'fact';
          }
        }
        // Unknown object kind
        else {
          console.warn(`Unknown object kind: ${obj.kind} for ${name}, semanticType: ${semanticType}`);
        }

      });

      // If still no fields, log sample objects for debugging
      if (dimensions.length === 0 && metrics.length === 0 && facts.length === 0) {
      }

      // ENHANCEMENT: Enrich dimensions and facts with table names from SHOW commands
      // DESC SEMANTIC VIEW doesn't reliably return table names, so we use SHOW commands
      try {
        // Get table names for dimensions
        const dimensionTableMap = await this.getTableNamesForDimensions(fullyQualifiedName);
        dimensions.forEach(dim => {
          if (dimensionTableMap.has(dim.name)) {
            dim.logicalTable = dimensionTableMap.get(dim.name);
          }
        });

        // Get table names for facts
        const factTableMap = await this.getTableNamesForFacts(fullyQualifiedName);
        facts.forEach(fact => {
          if (factTableMap.has(fact.name)) {
            fact.logicalTable = factTableMap.get(fact.name);
          }
        });
      } catch (enrichError) {
        console.warn('Failed to enrich metadata with table names:', enrichError.message);
        // Continue without table names - validation will be limited
      }

      return {
        dimensions,
        metrics,
        facts,
        nonEnforcedFilters
      };
    } catch (error) {
      throw new Error(`Failed to describe semantic view: ${error.message}`);
    }
  },

  /**
   * Get table names for all dimensions using SHOW SEMANTIC DIMENSIONS
   * Returns: Map<dimensionName, tableName>
   */
  async getTableNamesForDimensions(semanticView) {
    try {
      const sql = `SHOW SEMANTIC DIMENSIONS IN ${semanticView}`;
      const result = await this.executeQuery(sql);

      const tableMap = new Map();
      result.rows.forEach(row => {
        const dimName = row.name || row.NAME;
        const tableName = row.table_name || row.TABLE_NAME;
        if (dimName && tableName) {
          tableMap.set(dimName, tableName);
        }
      });

      return tableMap;
    } catch (error) {
      console.error('Failed to get dimension table names:', error);
      return new Map(); // Return empty map on error
    }
  },

  /**
   * Get table names for all facts using SHOW SEMANTIC FACTS
   * Returns: Map<factName, tableName>
   */
  async getTableNamesForFacts(semanticView) {
    try {
      const sql = `SHOW SEMANTIC FACTS IN ${semanticView}`;
      const result = await this.executeQuery(sql);

      const tableMap = new Map();
      result.rows.forEach(row => {
        const factName = row.name || row.NAME;
        const tableName = row.table_name || row.TABLE_NAME;
        if (factName && tableName) {
          tableMap.set(factName, tableName);
        }
      });

      return tableMap;
    } catch (error) {
      console.error('Failed to get fact table names:', error);
      return new Map(); // Return empty map on error
    }
  },

  /**
   * Get distinct values for a column (for filter dropdowns)
   */
  async getDistinctValues(semanticView, column, limit = QUERY_CONFIG.DEFAULT_DISTINCT_VALUES_LIMIT) {
    try {
      // First check count (excluding NULLs)
      const countSql = `SELECT COUNT(DISTINCT ${column}) as cnt FROM ${semanticView} WHERE ${column} IS NOT NULL`;
      const countResult = await this.executeQuery(countSql);
      const distinctCount = parseInt(countResult.rows[0].cnt, 10);

      if (distinctCount > limit) {
        throw new Error(
          `Too many distinct values (${distinctCount.toLocaleString()}). ` +
          `Limit is ${limit.toLocaleString()}. Please use a different filter.`
        );
      }

      // Get distinct values (excluding NULLs)
      const sql = `SELECT DISTINCT ${column} FROM ${semanticView} WHERE ${column} IS NOT NULL ORDER BY ${column} LIMIT ${limit}`;
      const result = await this.executeQuery(sql);

      // Filter out any remaining null/undefined values as a safety measure
      return result.rows.map(row => row[column]).filter(v => v !== null && v !== undefined);
    } catch (error) {
      throw new Error(`Failed to get distinct values: ${error.message}`);
    }
  },

  /**
   * Get allowed dimensions for a specific metric
   * Uses: SHOW SEMANTIC DIMENSIONS IN <view> FOR METRIC <metric_name>
   *
   * Returns dimensions that meet Snowflake's constraints:
   * 1. The dimension's logical table must be related to the metric's logical table
   * 2. The dimension must have equal or lower granularity than the metric
   *
   * This is the authoritative way to determine dimension-metric compatibility.
   */
  async getDimensionsForMetric(semanticView, metricName) {
    try {
      const sql = `SHOW SEMANTIC DIMENSIONS IN ${semanticView} FOR METRIC ${metricName}`;
      const result = await this.executeQuery(sql);

      // Parse results - returns dimensions that are compatible with this metric
      return result.rows.map(row => ({
        tableName: row.table_name,
        name: row.name,
        dataType: row.data_type,
        required: row.required === 'true' || row.required === true,
        fullyQualifiedName: `${row.table_name}.${row.name}`
      }));
    } catch (error) {
      console.error(`Failed to get dimensions for metric ${metricName}:`, error);
      return []; // Return empty array if command fails
    }
  },

  /**
   * Get primary keys for all logical tables in a semantic view
   * Uses: DESC SEMANTIC VIEW to extract PRIMARY_KEY property from TABLE objects
   * Returns: Map<tableName, primaryKeys[]>
   */
  async getTablePrimaryKeys(semanticView) {
    try {
      const sql = `DESC SEMANTIC VIEW ${semanticView}`;
      const result = await this.executeQuery(sql);

      const tableKeys = new Map();

      result.rows.forEach(row => {
        const objectKind = row.object_kind || row.OBJECT_KIND;
        const objectName = row.object_name || row.OBJECT_NAME;
        const property = row.property || row.PROPERTY;
        const propertyValue = row.property_value || row.PROPERTY_VALUE;

        // Look for TABLE objects with PRIMARY_KEY property
        if (objectKind === 'TABLE' && property === 'PRIMARY_KEY' && propertyValue) {
          // PRIMARY_KEY can be comma-separated list
          const primaryKeys = propertyValue.split(',').map(k => k.trim()).filter(k => k);
          tableKeys.set(objectName, primaryKeys);
        }
      });

      return tableKeys;
    } catch (error) {
      console.error('Failed to get table primary keys:', error);
      return new Map(); // Return empty map on error
    }
  },

  /**
   * Validate if selected dimensions can uniquely determine selected facts
   * Returns validation result with warnings/errors for non-deterministic combinations
   */
  async validateFactDimensionCombination(semanticView, selectedDimensions, selectedFacts, metadata) {
    try {
      // If no facts selected, validation passes
      if (!selectedFacts || selectedFacts.length === 0) {
        return { valid: true, warnings: [], errors: [] };
      }

      // Get primary keys for all tables
      const tableKeys = await this.getTablePrimaryKeys(semanticView);

      // Group facts by their logical table
      const factsByTable = new Map();
      selectedFacts.forEach(factName => {
        const fact = metadata.facts.find(f => f.name === factName);
        if (fact && fact.logicalTable) {
          if (!factsByTable.has(fact.logicalTable)) {
            factsByTable.set(fact.logicalTable, []);
          }
          factsByTable.get(fact.logicalTable).push(fact);
        }
      });

      // Group dimensions by their logical table
      const dimensionsByTable = new Map();
      selectedDimensions.forEach(dimName => {
        const dimension = metadata.dimensions.find(d => d.name === dimName);
        if (dimension && dimension.logicalTable) {
          if (!dimensionsByTable.has(dimension.logicalTable)) {
            dimensionsByTable.set(dimension.logicalTable, []);
          }
          dimensionsByTable.get(dimension.logicalTable).push(dimension);
        }
      });

      const warnings = [];
      const errors = [];
      const factTables = Array.from(factsByTable.keys());

      // Check if facts are from multiple tables
      if (factTables.length > 1) {
        errors.push({
          type: 'MULTIPLE_FACT_TABLES',
          message: `Facts are from multiple tables: ${factTables.join(', ')}. Facts must come from a single logical table.`,
          factTables: factTables
        });
      }

      // For each fact table, validate dimensions
      factsByTable.forEach((facts, tableName) => {
        const requiredKeys = tableKeys.get(tableName) || [];
        const dimsFromSameTable = dimensionsByTable.get(tableName) || [];

        // Check if we have dimensions from the same table as the facts
        if (dimsFromSameTable.length === 0) {
          errors.push({
            type: 'NO_DIMENSIONS_FROM_FACT_TABLE',
            table: tableName,
            facts: facts.map(f => f.name),
            message: `Facts from table '${tableName}' require dimensions from the same table. ` +
                     `Consider using METRICS instead of FACTS when combining with dimensions from other tables.`
          });
          return;
        }

        // Check if dimensions include primary keys
        const dimNamesFromTable = dimsFromSameTable.map(d => d.name);
        const missingKeys = requiredKeys.filter(key => !dimNamesFromTable.includes(key));

        if (missingKeys.length > 0 && requiredKeys.length > 0) {
          warnings.push({
            type: 'MISSING_PRIMARY_KEYS',
            table: tableName,
            facts: facts.map(f => f.name),
            requiredKeys: requiredKeys,
            missingKeys: missingKeys,
            message: `Query may be non-deterministic. Facts from '${tableName}' are missing primary key dimensions: ${missingKeys.join(', ')}. ` +
                     `Results will be grouped by selected dimensions, which may not uniquely identify each fact.`
          });
        }

        // Check if there are dimensions from OTHER tables
        const dimsFromOtherTables = selectedDimensions.filter(dimName => {
          const dim = metadata.dimensions.find(d => d.name === dimName);
          return dim && dim.logicalTable !== tableName;
        });

        if (dimsFromOtherTables.length > 0) {
          errors.push({
            type: 'DIMENSIONS_FROM_OTHER_TABLES',
            table: tableName,
            facts: facts.map(f => f.name),
            incompatibleDimensions: dimsFromOtherTables,
            message: `Cannot mix facts from '${tableName}' with dimensions from other tables: ${dimsFromOtherTables.join(', ')}. ` +
                     `Either: (1) Remove dimensions from other tables, (2) Remove the facts, or (3) Use METRICS instead of FACTS.`
          });
        }
      });

      return {
        valid: errors.length === 0,
        warnings: warnings,
        errors: errors
      };
    } catch (error) {
      return {
        valid: true, // Allow query on validation error
        warnings: [{
          type: 'VALIDATION_ERROR',
          message: `Could not validate combination: ${error.message}`
        }],
        errors: []
      };
    }
  },

  /**
   * Get compatible fields that can be selected based on current selection
   * This enables dynamic filtering in the UI
   *
   * Returns: {
   *   compatibleDimensions: string[],  // dimension names that can be added
   *   compatibleFacts: string[],       // fact names that can be added
   *   compatibleMetrics: string[],     // metric names (always all, metrics work with any dims)
   *   mode: 'mixed'|'facts'|'metrics', // current query mode
   *   blockedReasons: Map<fieldName, reason>  // why each field is blocked
   * }
   */
  async getCompatibleFields(currentSelection, metadata) {
    const {
      facts: selectedFacts = [],
      metrics: selectedMetrics = []
    } = currentSelection;

    const blockedReasons = new Map();

    // Determine current mode
    let mode = 'mixed';
    if (selectedFacts.length > 0 && selectedMetrics.length === 0) {
      mode = 'facts';
    } else if (selectedMetrics.length > 0 && selectedFacts.length === 0) {
      mode = 'metrics';
    }

    // RULE 1: Cannot mix facts and metrics
    if (mode === 'facts') {
      // In facts mode: block all metrics
      metadata.metrics.forEach(metric => {
        blockedReasons.set(metric.name, {
          reason: 'FACTS_METRICS_CONFLICT',
          message: 'Cannot mix FACTS and METRICS in the same query'
        });
      });
    } else if (mode === 'metrics') {
      // In metrics mode: block all facts
      metadata.facts.forEach(fact => {
        blockedReasons.set(fact.name, {
          reason: 'FACTS_METRICS_CONFLICT',
          message: 'Cannot mix FACTS and METRICS in the same query'
        });
      });
    }

    // RULE 2: When facts are selected, only allow dimensions from the same table
    if (selectedFacts.length > 0) {
      // Find which table(s) the selected facts are from
      const factTables = new Set();
      selectedFacts.forEach(factName => {
        const fact = metadata.facts.find(f => f.name === factName);
        if (fact && fact.logicalTable) {
          factTables.add(fact.logicalTable);
        }
      });

      // Block dimensions from other tables
      metadata.dimensions.forEach(dim => {
        if (dim.logicalTable && !factTables.has(dim.logicalTable)) {
          blockedReasons.set(dim.name, {
            reason: 'DIMENSION_DIFFERENT_TABLE',
            message: `When using facts from '${Array.from(factTables).join(', ')}', you can only add dimensions from the same table`,
            factTables: Array.from(factTables),
            dimensionTable: dim.logicalTable
          });
        }
      });

      // Block facts from other tables
      metadata.facts.forEach(fact => {
        if (fact.logicalTable && !factTables.has(fact.logicalTable)) {
          blockedReasons.set(fact.name, {
            reason: 'FACT_DIFFERENT_TABLE',
            message: `Facts must come from the same table: ${Array.from(factTables).join(', ')}`,
            factTables: Array.from(factTables),
            factTable: fact.logicalTable
          });
        }
      });
    }

    // Build compatible lists (fields not in blockedReasons)
    const compatibleDimensions = metadata.dimensions
      .filter(d => !blockedReasons.has(d.name))
      .map(d => d.name);

    const compatibleFacts = metadata.facts
      .filter(f => !blockedReasons.has(f.name))
      .map(f => f.name);

    const compatibleMetrics = metadata.metrics
      .filter(m => !blockedReasons.has(m.name))
      .map(m => m.name);

    return {
      compatibleDimensions,
      compatibleFacts,
      compatibleMetrics,
      mode,
      blockedReasons,
      // Helper method to check if a field is compatible
      isFieldCompatible: (fieldName) => !blockedReasons.has(fieldName),
      // Get reason why a field is blocked
      getBlockedReason: (fieldName) => blockedReasons.get(fieldName)
    };
  },

  /**
   * Get row count for a query
   */
  async getRowCount(query) {
    try {
      const countQuery = `SELECT COUNT(*) as cnt FROM (${query})`;
      const result = await this.executeQuery(countQuery);
      return parseInt(result.rows[0].cnt, 10);
    } catch (error) {
      throw new Error(`Failed to get row count: ${error.message}`);
    }
  },

  /**
   * Validate if a dimension can be used with a metric
   * Uses: SHOW SEMANTIC DIMENSION FOR METRIC
   *
   * TODO: Consider implementing caching for metadata calls (describeSemanticView,
   * validateDimensionForMetric, getDatabases, getSchemas, etc.) to improve performance
   * and reduce server load. Metadata doesn't change frequently within a session.
   */
  async validateDimensionForMetric(semanticView, metricName, dimensionName) {
    try {
      const sql = `SHOW SEMANTIC DIMENSION FOR METRIC ${metricName} IN ${semanticView}`;
      const result = await this.executeQuery(sql);

      const allowedDimensions = result.rows.map(row => row.dimension_name);

      if (!allowedDimensions.includes(dimensionName)) {
        return {
          valid: false,
          error: `Dimension "${dimensionName}" cannot be used with metric "${metricName}". ` +
                 `Allowed dimensions: ${allowedDimensions.join(', ')}`
        };
      }

      return { valid: true };
    } catch (error) {
      // If validation command not supported, assume valid
      return { valid: true, warning: 'Validation not available' };
    }
  },

  /**
   * Helper: Check if a data type is numeric
   */
  isNumericType(type) {
    const numericTypes = [
      'NUMBER', 'INT', 'INTEGER', 'BIGINT', 'SMALLINT', 'TINYINT',
      'FLOAT', 'DOUBLE', 'DECIMAL', 'NUMERIC', 'REAL'
    ];
    return numericTypes.some(t => type.toUpperCase().includes(t));
  },

  /**
   * Execute a semantic query with adhoc metrics
   */
  async executeSemanticQuery(config) {
    const {
      semanticView,
      dimensions = [],
      metrics = [],
      facts = [],
      adhocMetrics = [],
      filters = [],
      limit = QUERY_CONFIG.DEFAULT_RESULT_LIMIT,
      offset = 0
    } = config;

    try {
      // Validate: Must have at least one field
      if (dimensions.length === 0 && metrics.length === 0 && facts.length === 0 && adhocMetrics.length === 0) {
        throw new Error('Please select at least one field (dimension, metric, or fact)');
      }

      // Validate: Cannot mix FACTS and METRICS
      if (facts.length > 0 && (metrics.length > 0 || adhocMetrics.length > 0)) {
        throw new Error('Cannot specify both FACTS and METRICS in the same query. Please use either facts or metrics, not both.');
      }

      // Build SEMANTIC_VIEW clauses
      // NOTE: Clause order per Snowflake docs:
      // 1. METRICS or FACTS (mutually exclusive)
      // 2. DIMENSIONS
      // 3. WHERE
      const clauses = [];

      // METRICS clause (includes adhoc metrics) - comes FIRST
      if (metrics.length > 0 || adhocMetrics.length > 0) {
        const allMetrics = [...metrics];

        // Add adhoc metrics with their expressions
        adhocMetrics.forEach(adhoc => {
          allMetrics.push(`${adhoc.expression} AS ${adhoc.name}`);
        });

        clauses.push(`METRICS ${allMetrics.join(', ')}`);
      }

      // FACTS clause - comes FIRST (mutually exclusive with METRICS)
      if (facts.length > 0) {
        clauses.push(`FACTS ${facts.join(', ')}`);
      }

      // DIMENSIONS clause - comes AFTER METRICS/FACTS
      if (dimensions.length > 0) {
        clauses.push(`DIMENSIONS ${dimensions.join(', ')}`);
      }

      // Build WHERE clause
      if (filters.length > 0) {
        const conditions = filters.map(filter => {
          if (filter.operator === 'IN') {
            const values = filter.values.map(v => this.formatSQLValue(v)).join(', ');
            return `${filter.field} IN (${values})`;
          } else if (filter.operator === 'BETWEEN') {
            // BETWEEN requires exactly 2 values
            if (filter.values.length >= 2) {
              const min = this.formatSQLValue(filter.values[0]);
              const max = this.formatSQLValue(filter.values[1]);
              return `${filter.field} BETWEEN ${min} AND ${max}`;
            }
            return '';
          } else if (['=', '!=', '>', '<', '>=', '<='].includes(filter.operator)) {
            return `${filter.field} ${filter.operator} ${this.formatSQLValue(filter.values[0])}`;
          } else if (filter.operator === 'LIKE') {
            // LIKE always needs quoted strings
            return `${filter.field} LIKE '${this.escapeSQL(filter.values[0])}'`;
          }
          return '';
        }).filter(Boolean);

        if (conditions.length > 0) {
          clauses.push(`WHERE ${conditions.join(' AND ')}`);
        }
      }

      // Build final SEMANTIC_VIEW query
      const semanticViewClause = clauses.join('\n    ');

      let query = `SELECT * FROM SEMANTIC_VIEW(
    ${semanticView}
    ${semanticViewClause}
  )`;

      // Add LIMIT and OFFSET outside the SEMANTIC_VIEW clause
      query += ` LIMIT ${limit} OFFSET ${offset}`;


      // Store the query for error handling
      this._lastQuery = query;

      const result = await this.executeQuery(query);

      // Return both the result and the SQL query
      result.sql = query;
      return result;
    } catch (error) {
      // Create error with SQL attached
      const err = new Error(`Failed to execute semantic query: ${error.message}`);
      err.sql = this._lastQuery;
      throw err;
    }
  },

  /**
   * Helper: Check if a value should be treated as numeric (not quoted)
   */
  isNumericValue(value) {
    // Check if value is a number or a string that represents a valid number
    if (typeof value === 'number') return true;
    if (typeof value === 'string') {
      const trimmed = value.trim();
      // Check if it's a valid number (including decimals, negative numbers)
      return /^-?\d+(\.\d+)?$/.test(trimmed);
    }
    return false;
  },

  /**
   * Helper: Format SQL value (quote strings, leave numbers unquoted)
   */
  formatSQLValue(value) {
    if (this.isNumericValue(value)) {
      return value; // Don't quote numbers
    }
    // Quote and escape strings
    return `'${this.escapeSQL(value)}'`;
  },

  /**
   * Helper: Escape SQL string values
   */
  escapeSQL(value) {
    if (typeof value === 'string') {
      return value.replace(/'/g, "''");
    }
    return value;
  }
};

export default SnowflakeAPI;
