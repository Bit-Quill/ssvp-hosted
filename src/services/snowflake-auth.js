import PKCEHelper from '../utils/pkce-helper.js';

const AUTH_CONFIG = {
  DIALOG_HEIGHT_PERCENT: 60,
  DIALOG_WIDTH_PERCENT: 30,
  POPUP_WIDTH_PX: 500,
  POPUP_HEIGHT_PX: 600,
  POLLING_INTERVAL_MS: 1000,
  OAUTH_TIMEOUT_MS: 10 * 60 * 1000,
  POPUP_TIMEOUT_MS: 5 * 60 * 1000,
  TOKEN_EXPIRY_BUFFER_MS: 5 * 60 * 1000
};

const SnowflakeAuth = {
  config: {
    account: null,
    username: null,
    warehouse: null,
    role: null,
    authMethod: 'oauth',
    clientId: null,
    // Note: clientSecret should NEVER be stored client-side
    // It should only exist on the backend server
    integrationName: null,
    redirectUri: null, // Will be set dynamically based on clientId
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
      // Default to LOCAL_APPLICATION for simplest setup
      this.config.clientId = 'LOCAL_APPLICATION';
      this.config.integrationName = 'SNOWFLAKE$LOCAL_APPLICATION';
      this.config.redirectUri = 'http://127.0.0.1';
    } else {
      // For custom OAuth clients, use the configured redirect URI
      // This is injected at build time by webpack based on deployment environment
      if (!this.config.redirectUri) {
        this.config.redirectUri = process.env.OAUTH_REDIRECT_URI || 'http://localhost:3000/auth/callback';
      }
    }
  },

  async loginWithOAuthDialog() {
    try {
      const codeVerifier = PKCEHelper.generateCodeVerifier();
      const codeChallenge = await PKCEHelper.generateCodeChallenge(codeVerifier);
      const state = PKCEHelper.generateState();

      PKCEHelper.storePKCEParams(codeVerifier, state);

      const authUrl = this.buildOAuthUrl(codeChallenge, state);
      const dialogUrl = `${window.location.origin}/auth/auth-start.html?authUrl=${encodeURIComponent(authUrl)}`;

      return new Promise((resolve, reject) => {
        Office.context.ui.displayDialogAsync(
          dialogUrl,
          { height: AUTH_CONFIG.DIALOG_HEIGHT_PERCENT, width: AUTH_CONFIG.DIALOG_WIDTH_PERCENT, displayInIframe: false },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              PKCEHelper.clearPKCEParams();
              return reject(new Error(`Failed to open dialog: ${asyncResult.error.message}`));
            }

            const dialog = asyncResult.value;

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
              dialog.close();

              try {
                const messageData = JSON.parse(arg.message);

                if (messageData.type === 'SNOWFLAKE_AUTH_SUCCESS') {
                  if (!PKCEHelper.validateState(messageData.state)) {
                    PKCEHelper.clearPKCEParams();
                    return reject(new Error('Invalid state parameter'));
                  }

                  const result = await this._completeOAuthFlow(messageData.code, codeVerifier);
                  resolve(result);
                } else if (messageData.type === 'SNOWFLAKE_AUTH_ERROR') {
                  PKCEHelper.clearPKCEParams();
                  reject(new Error(messageData.error || 'Authentication failed'));
                }
              } catch (error) {
                PKCEHelper.clearPKCEParams();
                reject(error);
              }
            });

            const pollInterval = setInterval(() => {
              try {
                const result = localStorage.getItem('snowflake_oauth_result');
                if (result) {
                  clearInterval(pollInterval);
                  localStorage.removeItem('snowflake_oauth_result');
                  dialog.close();

                  const authData = JSON.parse(result);

                  if (!PKCEHelper.validateState(authData.state)) {
                    PKCEHelper.clearPKCEParams();
                    return reject(new Error('Invalid state parameter'));
                  }

                  this._completeOAuthFlow(authData.code, codeVerifier)
                    .then(result => resolve(result))
                    .catch(error => reject(error));
                }
              } catch (error) {
                console.error('localStorage polling error:', error);
              }
            }, 1000);

            dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
              clearInterval(pollInterval);
              if (arg.error === 12006) {
                PKCEHelper.clearPKCEParams();
                reject(new Error('Authentication cancelled'));
              }
            });
          }
        );
      });
    } catch (error) {
      PKCEHelper.clearPKCEParams();
      this.logout();
      throw new Error(`OAuth login failed: ${error.message}`);
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
        throw new Error('Invalid state parameter');
      }

      const result = await this._completeOAuthFlow(authResult.code, codeVerifier);

      return {
        ...result,
        warehouse: this.config.warehouse,
        role: this.config.role
      };
    } catch (error) {
      PKCEHelper.clearPKCEParams();
      this.logout();
      throw new Error(`OAuth login failed: ${error.message}`);
    }
  },

  _setAuthTokens(tokens) {
    this.accessToken = tokens.access_token;
    this.refreshToken = tokens.refresh_token;
    this.tokenExpiry = Date.now() + (tokens.expires_in * 1000);
    this.authMethod = 'oauth';
  },

  async _completeOAuthFlow(code, codeVerifier) {
    const tokens = await this.exchangeCodeForTokens(code, codeVerifier);
    this._setAuthTokens(tokens);
    PKCEHelper.clearPKCEParams();
    return {
      success: true,
      token: this.accessToken,
      expiresIn: tokens.expires_in
    };
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
    const width = AUTH_CONFIG.POPUP_WIDTH_PX;
    const height = AUTH_CONFIG.POPUP_HEIGHT_PX;
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

  waitForManualAuth() {
    return new Promise((resolve, reject) => {
      const startTime = Date.now();
      const timeout = 600000;

      const pollInterval = setInterval(() => {
        try {
          const result = localStorage.getItem('snowflake_oauth_result');

          if (result) {
            clearInterval(pollInterval);

            const authData = JSON.parse(result);
            localStorage.removeItem('snowflake_oauth_result');

            if (!PKCEHelper.validateState(authData.state)) {
              PKCEHelper.clearPKCEParams();
              return reject(new Error('Invalid state parameter - possible CSRF attack'));
            }

            this.exchangeCodeForTokens(authData.code, PKCEHelper.retrieveCodeVerifier())
              .then(tokens => {
                this.accessToken = tokens.access_token;
                this.refreshToken = tokens.refresh_token;
                this.tokenExpiry = Date.now() + (tokens.expires_in * 1000);
                this.authMethod = 'oauth';

                PKCEHelper.clearPKCEParams();

                resolve({
                  success: true,
                  token: this.accessToken,
                  expiresIn: tokens.expires_in
                });
              })
              .catch(error => {
                PKCEHelper.clearPKCEParams();
                reject(error);
              });
          }

          if (Date.now() - startTime > timeout) {
            clearInterval(pollInterval);
            PKCEHelper.clearPKCEParams();
            reject(new Error('Authentication timeout - please try again'));
          }
        } catch (error) {
          clearInterval(pollInterval);
          PKCEHelper.clearPKCEParams();
          reject(error);
        }
      }, 1000);
    });
  },

  waitForAuthCallback(authWindow) {
    return new Promise((resolve, reject) => {
      const messageHandler = (event) => {
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

          try {
            const result = localStorage.getItem('snowflake_oauth_result');
            if (result) {
              localStorage.removeItem('snowflake_oauth_result');
              const authData = JSON.parse(result);
              return resolve({
                code: authData.code,
                state: authData.state
              });
            }
          } catch (err) {
          }

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
      // For custom OAuth clients (not LOCAL_APPLICATION), call Snowflake directly
      // This eliminates the need for a backend server
      const tokenUrl = `https://${this.config.account}.snowflakecomputing.com/oauth/token-request`;

      const tokenParams = new URLSearchParams({
        grant_type: 'authorization_code',
        code: code,
        redirect_uri: this.config.redirectUri,
        code_verifier: codeVerifier,
        client_id: this.config.clientId
      });

      // For LOCAL_APPLICATION, add the client_secret
      if (this.config.clientId === 'LOCAL_APPLICATION') {
        tokenParams.append('client_secret', 'LOCAL_APPLICATION');
      }
      // For public OAuth clients, no client_secret is needed (PKCE provides security)

      const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Accept': 'application/json'
        },
        body: tokenParams.toString()
      });

      let result;
      try {
        result = await response.json();
      } catch (parseError) {
        const text = await response.text();
        throw new Error(`Invalid response from Snowflake: ${text}`);
      }

      if (!response.ok) {
        const errorMsg = result.error_description || result.error || 'Token exchange failed';
        throw new Error(errorMsg);
      }

      if (!result.access_token) {
        throw new Error('No access token received from Snowflake');
      }

      return result;
    } catch (error) {
      throw new Error(`Token exchange failed: ${error.message}`);
    }
  },

  async refreshAccessToken() {
    if (!this.refreshToken) {
      throw new Error('No refresh token available');
    }

    try {
      // Call Snowflake directly - no backend needed
      const tokenUrl = `https://${this.config.account}.snowflakecomputing.com/oauth/token-request`;

      const tokenParams = new URLSearchParams({
        grant_type: 'refresh_token',
        refresh_token: this.refreshToken,
        client_id: this.config.clientId
      });

      // For LOCAL_APPLICATION, add the client_secret
      if (this.config.clientId === 'LOCAL_APPLICATION') {
        tokenParams.append('client_secret', 'LOCAL_APPLICATION');
      }
      // For public OAuth clients, no client_secret is needed

      const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Accept': 'application/json'
        },
        body: tokenParams.toString()
      });

      let result;
      try {
        result = await response.json();
      } catch (parseError) {
        const text = await response.text();
        throw new Error(`Invalid response from Snowflake: ${text}`);
      }

      if (!response.ok) {
        const errorMsg = result.error_description || result.error || 'Token refresh failed';
        throw new Error(errorMsg);
      }

      this.accessToken = result.access_token;
      if (result.refresh_token) {
        this.refreshToken = result.refresh_token;
      }
      this.tokenExpiry = Date.now() + (result.expires_in * 1000);

      return result;
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
