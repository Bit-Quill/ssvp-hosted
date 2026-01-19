/**
 * Snowflake API Service
 * Provides methods for interacting with Snowflake Semantic Views
 */

import SnowflakeAuth from './snowflake-auth.js';

const SnowflakeAPI = {
  /**
   * Execute a SQL query via Snowflake SQL API
   */
  async executeQuery(sql, options = {}) {
    try {
      const token = await SnowflakeAuth.getAccessToken();
      const { timeout = 60, warehouse = null, role = null } = options;

      const requestBody = {
        statement: sql,
        timeout: timeout,
        resultSetMetaData: { format: 'json' }
      };

      if (warehouse || SnowflakeAuth.config.warehouse) {
        requestBody.warehouse = warehouse || SnowflakeAuth.config.warehouse;
      }

      if (role || SnowflakeAuth.config.role) {
        requestBody.role = role || SnowflakeAuth.config.role;
      }

      const response = await fetch(
        `https://${SnowflakeAuth.config.account}.snowflakecomputing.com/api/v2/statements`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
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

      // Debug logging
      console.log('Query result:', {
        hasMetadata: !!result.resultSetMetaData,
        hasData: !!result.data,
        hasStatusUrl: !!result.statementStatusUrl,
        hasHandle: !!result.statementHandle,
        message: result.message
      });

      // Check for immediate success
      if (result.resultSetMetaData && result.data) {
        return this.parseResults(result);
      }

      // Handle async execution (query still running)
      if (result.statementStatusUrl || result.statementHandle) {
        const statusUrl = result.statementStatusUrl ||
          `https://${SnowflakeAuth.config.account}.snowflakecomputing.com/api/v2/statements/${result.statementHandle}`;
        console.log('Polling status URL:', statusUrl);
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
    const maxAttempts = 60;
    let attempts = 0;
    const initialDelay = 500; // Start with 500ms
    const maxDelay = 5000;     // Cap at 5 seconds
    const backoffMultiplier = 1.5;

    while (attempts < maxAttempts) {
      try {
        const token = await SnowflakeAuth.getAccessToken();
        const response = await fetch(statusUrl, {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          }
        });

        if (!response.ok) {
          const errorText = await response.text();
          console.error('Poll failed:', response.status, errorText);
          throw new Error(`Failed to poll query status: ${response.status} ${errorText}`);
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
        const delay = Math.min(initialDelay * Math.pow(backoffMultiplier, attempts), maxDelay);
        await new Promise(resolve => setTimeout(resolve, delay));
        attempts++;
      } catch (error) {
        console.error('Polling error:', error);
        throw error;
      }
    }

    throw new Error('Query timed out after 60 seconds');
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
   * Implements: Metadata Parsing requirement
   * Returns: { dimensions, metrics, facts, nonEnforcedFilters }
   */
  async describeSemanticView(fullyQualifiedName) {
    try {
      const sql = `DESC SEMANTIC VIEW ${fullyQualifiedName}`;
      const result = await this.executeQuery(sql);

      const dimensions = [];
      const metrics = [];
      const facts = [];
      const nonEnforcedFilters = [];

      result.rows.forEach(row => {
        const field = {
          name: row.name,
          dataType: row.type,
          description: row.description || '',
          kind: row.kind
        };

        switch (row.kind) {
          case 'DIMENSION':
            dimensions.push({
              ...field,
              logicalTable: row.logical_table || null
            });
            break;

          case 'METRIC':
            metrics.push({
              ...field,
              expression: row.expr || null,
              logicalTable: row.logical_table || null
            });
            break;

          case 'FACT':
            const isNumeric = this.isNumericType(row.type);
            facts.push({
              ...field,
              isNumeric: isNumeric,
              canBeMetric: isNumeric,
              logicalTable: row.logical_table || null
            });
            break;
        }
      });

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
   * Get distinct values for a column (for filter dropdowns)
   */
  async getDistinctValues(semanticView, column, limit = 1000) {
    try {
      // First check count
      const countSql = `SELECT COUNT(DISTINCT ${column}) as cnt FROM ${semanticView}`;
      const countResult = await this.executeQuery(countSql);
      const distinctCount = parseInt(countResult.rows[0].cnt, 10);

      if (distinctCount > limit) {
        throw new Error(
          `Too many distinct values (${distinctCount.toLocaleString()}). ` +
          `Limit is ${limit.toLocaleString()}. Please use a different filter.`
        );
      }

      // Get distinct values
      const sql = `SELECT DISTINCT ${column} FROM ${semanticView} ORDER BY ${column} LIMIT ${limit}`;
      const result = await this.executeQuery(sql);

      return result.rows.map(row => row[column]);
    } catch (error) {
      throw new Error(`Failed to get distinct values: ${error.message}`);
    }
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
      adhocMetrics = [],
      filters = [],
      limit = 10000,
      offset = 0
    } = config;

    try {
      // Must have at least one field selected
      if (dimensions.length === 0 && metrics.length === 0 && adhocMetrics.length === 0) {
        throw new Error('Please select at least one field');
      }

      // Build DIMENSIONS clause
      let dimensionsClause = '';
      if (dimensions.length > 0) {
        dimensionsClause = `DIMENSIONS ${dimensions.join(', ')}`;
      }

      // Build METRICS clause (includes both predefined metrics and adhoc metrics)
      let metricsClause = '';
      if (metrics.length > 0 || adhocMetrics.length > 0) {
        const allMetrics = [...metrics];

        // Add adhoc metrics with their expressions
        adhocMetrics.forEach(adhoc => {
          allMetrics.push(`${adhoc.expression} AS ${adhoc.name}`);
        });

        metricsClause = `METRICS ${allMetrics.join(', ')}`;
      }

      // Build WHERE clause
      let whereClause = '';
      if (filters.length > 0) {
        const conditions = filters.map(filter => {
          if (filter.operator === 'IN') {
            const values = filter.values.map(v => `'${this.escapeSQL(v)}'`).join(', ');
            return `${filter.field} IN (${values})`;
          } else if (['=', '>', '<'].includes(filter.operator)) {
            return `${filter.field} ${filter.operator} '${this.escapeSQL(filter.values[0])}'`;
          }
          return '';
        }).filter(Boolean);

        if (conditions.length > 0) {
          whereClause = `WHERE ${conditions.join(' AND ')}`;
        }
      }

      // Build SEMANTIC_VIEW() query
      // Note: GROUP BY is not needed - semantic views handle aggregation automatically
      const semanticViewClauses = [
        semanticView,
        metricsClause,
        dimensionsClause,
        whereClause
      ].filter(Boolean).join('\n  ');

      const query = `SELECT * FROM SEMANTIC_VIEW(\n  ${semanticViewClauses}\n) LIMIT ${limit} OFFSET ${offset}`;

      const result = await this.executeQuery(query);
      return result;
    } catch (error) {
      throw new Error(`Failed to execute semantic query: ${error.message}`);
    }
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
