import PKCEHelper from '../utils/pkce-helper.js';

const SnowflakeAuth = {
  config: {
    account: null,
    username: null,
    warehouse: null,
    role: null,
    authMethod: 'oauth',
    clientId: null,
    clientSecret: null,
    integrationName: null,
    redirectUri: 'https://localhost:3000/auth/callback',
    scopes: ['refresh_token']
  },

  accessToken: null,
  refreshToken: null,
  tokenExpiry: null,
  authMethod: null,
  authWindow: null,

  init(config) {
    this.config = { ...this.config, ...config };

    if (!this.config.account) {
      throw new Error('Snowflake account identifier is required');
    }

    if (!this.config.clientId) {
      this.config.clientId = 'LOCAL_APPLICATION';
      this.config.integrationName = 'SNOWFLAKE$LOCAL_APPLICATION';
    }
  },

  async loginWithOAuth() {
    const authWindow = this.openAuthWindowEarly();

    if (!authWindow) {
      throw new Error('Failed to open authorization window. Please allow popups for this add-in.');
    }

    try {
      const codeVerifier = PKCEHelper.generateCodeVerifier();
      const codeChallenge = await PKCEHelper.generateCodeChallenge(codeVerifier);
      const state = PKCEHelper.generateState();

      PKCEHelper.storePKCEParams(codeVerifier, state);

      const authUrl = this.buildOAuthUrl(codeChallenge, state);
      authWindow.location.href = authUrl;

      const authResult = await this.waitForAuthCallback(authWindow);

      if (!PKCEHelper.validateState(authResult.state)) {
        throw new Error('Invalid state parameter - possible CSRF attack');
      }

      const tokens = await this.exchangeCodeForTokens(authResult.code, codeVerifier);

      this.accessToken = tokens.access_token;
      this.refreshToken = tokens.refresh_token;
      this.tokenExpiry = Date.now() + (tokens.expires_in * 1000);
      this.authMethod = 'oauth';

      PKCEHelper.clearPKCEParams();

      return {
        success: true,
        token: this.accessToken,
        warehouse: this.config.warehouse,
        role: this.config.role,
        expiresIn: tokens.expires_in
      };
    } catch (error) {
      PKCEHelper.clearPKCEParams();
      this.logout();
      throw new Error(`OAuth login failed: ${error.message}`);
    }
  },

  buildOAuthUrl(codeChallenge, state) {
    let scope = this.config.scopes.join(' ');

    if (this.config.role) {
      scope = `${scope} session:role:${this.config.role}`;
    }

    const params = new URLSearchParams({
      client_id: this.config.clientId,
      redirect_uri: this.config.redirectUri,
      response_type: 'code',
      scope: scope,
      state: state,
      code_challenge: codeChallenge,
      code_challenge_method: 'S256'
    });

    return `https://${this.config.account}.snowflakecomputing.com/oauth/authorize?${params.toString()}`;
  },

  openAuthWindowEarly() {
    const width = 500;
    const height = 600;
    const left = (window.screen.width - width) / 2;
    const top = (window.screen.height - height) / 2;

    const loadingUrl = `${window.location.origin}/auth/loading.html`;

    this.authWindow = window.open(
      loadingUrl,
      'snowflake-oauth',
      `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`
    );

    return this.authWindow;
  },

  waitForAuthCallback(authWindow) {
    return new Promise((resolve, reject) => {
      const messageHandler = (event) => {
        const expectedOrigin = window.location.origin;
        if (event.origin !== expectedOrigin) {
          return;
        }

        if (event.data.type === 'SNOWFLAKE_AUTH_SUCCESS') {
          window.removeEventListener('message', messageHandler);
          clearInterval(checkWindowClosed);
          clearTimeout(timeoutId);

          if (authWindow && !authWindow.closed) {
            authWindow.close();
          }

          resolve({
            code: event.data.code,
            state: event.data.state
          });
        } else if (event.data.type === 'SNOWFLAKE_AUTH_ERROR') {
          window.removeEventListener('message', messageHandler);
          clearInterval(checkWindowClosed);
          clearTimeout(timeoutId);

          if (authWindow && !authWindow.closed) {
            authWindow.close();
          }

          reject(new Error(event.data.error || 'Authentication failed'));
        }
      };

      window.addEventListener('message', messageHandler);

      const checkWindowClosed = setInterval(() => {
        if (authWindow && authWindow.closed) {
          clearInterval(checkWindowClosed);
          clearTimeout(timeoutId);
          window.removeEventListener('message', messageHandler);
          reject(new Error('Authentication window was closed'));
        }
      }, 1000);

      const timeoutId = setTimeout(() => {
        clearInterval(checkWindowClosed);
        window.removeEventListener('message', messageHandler);

        if (authWindow && !authWindow.closed) {
          authWindow.close();
        }

        reject(new Error('Authentication timeout'));
      }, 300000);
    });
  },

  async exchangeCodeForTokens(code, codeVerifier) {
    try {
      const proxyUrl = '/api/auth/token';

      const requestBody = {
        code: code,
        code_verifier: codeVerifier,
        redirect_uri: this.config.redirectUri,
        account: this.config.account,
        client_id: this.config.clientId
      };

      const response = await fetch(proxyUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
      });

      let result;
      try {
        result = await response.json();
      } catch (parseError) {
        const text = await response.text();
        throw new Error(`Invalid response from proxy: ${text}`);
      }

      if (!response.ok) {
        const errorMsg = result.message || result.error || 'Token exchange failed';
        const errorDetails = result.details ? JSON.stringify(result.details) : '';
        throw new Error(`${errorMsg}${errorDetails ? ' - ' + errorDetails : ''}`);
      }

      if (!result.access_token) {
        throw new Error('No access token received from Snowflake');
      }

      return result;
    } catch (error) {
      console.error('[SnowflakeAuth] Token exchange error:', {
        error: error.message,
        account: this.config.account,
        clientId: this.config.clientId,
        redirectUri: this.config.redirectUri
      });
      throw new Error(`Token exchange failed: ${error.message}`);
    }
  },

  async refreshAccessToken() {
    if (!this.refreshToken) {
      throw new Error('No refresh token available');
    }

    try {
      const response = await fetch('/api/auth/refresh', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          refresh_token: this.refreshToken,
          account: this.config.account
        })
      });

      if (!response.ok) {
        const error = await response.json();
        throw new Error(error.message || 'Token refresh failed');
      }

      const tokens = await response.json();

      this.accessToken = tokens.access_token;
      if (tokens.refresh_token) {
        this.refreshToken = tokens.refresh_token;
      }
      this.tokenExpiry = Date.now() + (tokens.expires_in * 1000);

      return tokens;
    } catch (error) {
      this.logout();
      throw new Error(`Token refresh failed: ${error.message}`);
    }
  },

  isTokenExpired() {
    if (!this.tokenExpiry) {
      return true;
    }

    const bufferTime = 5 * 60 * 1000;
    return Date.now() >= (this.tokenExpiry - bufferTime);
  },

  async getAccessToken() {
    if (!this.accessToken) {
      throw new Error('Not authenticated');
    }

    if (this.isTokenExpired() && this.refreshToken) {
      await this.refreshAccessToken();
    }

    return this.accessToken;
  },

  async loginWithKeyPair(username, privateKey) {
    throw new Error('Key-pair authentication not yet implemented');
  },

  async loginWithPAT(pat) {
    try {
      const response = await fetch(`https://${this.config.account}.snowflakecomputing.com/api/v2/statements`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${pat}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          statement: 'SELECT CURRENT_USER(), CURRENT_ROLE(), CURRENT_WAREHOUSE()',
          timeout: 10
        })
      });

      if (!response.ok) {
        throw new Error('Invalid Personal Access Token');
      }

      const result = await response.json();

      this.accessToken = pat;
      this.authMethod = 'pat';

      return {
        success: true,
        token: pat,
        warning: 'PAT authentication may not support Cortex Analyst features',
        user: result.data?.[0]?.[0],
        role: result.data?.[0]?.[1],
        warehouse: result.data?.[0]?.[2]
      };
    } catch (error) {
      throw new Error(`PAT authentication error: ${error.message}`);
    }
  },

  async executeStatement(sql, options = {}) {
    const token = await this.getAccessToken();

    const requestBody = {
      statement: sql,
      timeout: 60
    };

    if (options.warehouse || this.config.warehouse) {
      requestBody.warehouse = options.warehouse || this.config.warehouse;
    }

    if (options.role || this.config.role) {
      requestBody.role = options.role || this.config.role;
    }

    const response = await fetch(`https://${this.config.account}.snowflakecomputing.com/api/v2/statements`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.message || 'Statement execution failed');
    }

    return response.json();
  },

  setContext(warehouse, role) {
    if (warehouse) {
      this.config.warehouse = warehouse;
    }
    if (role) {
      this.config.role = role;
    }
  },

  logout() {
    this.accessToken = null;
    this.refreshToken = null;
    this.tokenExpiry = null;
    this.authMethod = null;

    PKCEHelper.clearPKCEParams();
    sessionStorage.clear();
  },

  isAuthenticated() {
    return !!this.accessToken && !this.isTokenExpired();
  },

  getAuthMethod() {
    return this.authMethod;
  }
};

export default SnowflakeAuth;
