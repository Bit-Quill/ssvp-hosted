const express = require('express');
const router = express.Router();

// SECURITY: Sanitize error responses to prevent information disclosure
// Logs full error details server-side but returns safe messages to client
function sanitizeError(error, context = '') {
  // Generate unique error ID for correlation between logs and client errors
  const errorId = `ERR-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

  // Log full error details server-side for debugging
  console.error(`[${errorId}] ${context}:`, {
    message: error.message || error,
    stack: error.stack,
    details: error.details || error,
    timestamp: new Date().toISOString()
  });

  // Return sanitized error to client with correlation ID
  return {
    errorId: errorId,
    message: getSafeErrorMessage(error, context)
  };
}

// Map internal errors to safe, user-friendly messages
function getSafeErrorMessage(error, context) {
  // Check for common error patterns and return appropriate user-friendly messages
  const errorString = typeof error === 'string' ? error : (error.message || '');

  // Snowflake-specific errors
  if (errorString.includes('invalid_grant') || errorString.includes('Invalid authorization code')) {
    return 'Authentication failed. Please try logging in again.';
  }
  if (errorString.includes('invalid_client')) {
    return 'OAuth client configuration error. Please contact support.';
  }
  if (errorString.includes('invalid_token') || errorString.includes('token')) {
    return 'Session expired. Please log in again.';
  }
  if (errorString.includes('access_denied')) {
    return 'Access denied. Please check your permissions.';
  }
  if (errorString.includes('Incorrect username or password') ||
      errorString.includes('Invalid username or password')) {
    return 'Incorrect username or password. Please try again.';
  }

  // Network/connectivity errors
  if (errorString.includes('ECONNREFUSED') || errorString.includes('ETIMEDOUT')) {
    return 'Unable to connect to Snowflake. Please check your network connection.';
  }

  // Default safe message based on context
  switch (context) {
    case 'token_exchange':
      return 'Failed to complete authentication. Please try again.';
    case 'token_refresh':
      return 'Failed to refresh session. Please log in again.';
    case 'token_revoke':
      return 'Logout completed. Session may not have been fully revoked.';
    default:
      return 'An error occurred. Please try again or contact support.';
  }
}

// SECURITY: Validate Snowflake account identifier to prevent SSRF attacks
// Account identifiers must match Snowflake's format specification
function validateSnowflakeAccount(account) {
  if (!account || typeof account !== 'string') {
    return { valid: false, error: 'Account identifier is required' };
  }

  // Snowflake account identifier format:
  // - Organization accounts: ORGNAME-ACCOUNTNAME (e.g., myorg-account123)
  // - Legacy accounts: ACCOUNTNAME (e.g., xy12345)
  // - With region: ACCOUNTNAME.REGION.CLOUD (e.g., xy12345.us-east-1.aws)
  // Valid characters: lowercase letters, numbers, hyphens, periods
  // Length: typically 2-255 characters
  const accountRegex = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)*$/i;

  if (!accountRegex.test(account)) {
    return {
      valid: false,
      error: 'Invalid Snowflake account identifier format. Must contain only letters, numbers, hyphens, and periods.'
    };
  }

  // Prevent excessively long account identifiers (potential DoS)
  if (account.length > 255) {
    return {
      valid: false,
      error: 'Account identifier exceeds maximum length'
    };
  }

  // Block localhost and internal network identifiers
  const blockedPatterns = [
    /^localhost$/i,
    /^127\./,
    /^10\./,
    /^172\.(1[6-9]|2[0-9]|3[01])\./,
    /^192\.168\./,
    /^0\.0\.0\.0$/,
    /^::1$/,
    /^fe80:/i
  ];

  for (const pattern of blockedPatterns) {
    if (pattern.test(account)) {
      return {
        valid: false,
        error: 'Account identifier contains blocked pattern'
      };
    }
  }

  // Optional: Check against allowlist if configured
  // This provides an additional layer of security for production deployments
  if (process.env.ALLOWED_SNOWFLAKE_ACCOUNTS) {
    const allowedAccounts = process.env.ALLOWED_SNOWFLAKE_ACCOUNTS.split(',').map(a => a.trim());
    if (!allowedAccounts.includes(account)) {
      return {
        valid: false,
        error: 'Account identifier not in allowlist'
      };
    }
  }

  return { valid: true };
}

router.post('/token', async (req, res) => {
  try {
    const { code, code_verifier, redirect_uri, account, client_id } = req.body;

    if (!code) {
      return res.status(400).json({
        success: false,
        message: 'Missing authorization code'
      });
    }

    if (!code_verifier) {
      return res.status(400).json({
        success: false,
        message: 'Missing code verifier (PKCE)'
      });
    }

    if (!redirect_uri) {
      return res.status(400).json({
        success: false,
        message: 'Missing redirect URI'
      });
    }

    if (!account) {
      return res.status(400).json({
        success: false,
        message: 'Missing Snowflake account'
      });
    }

    // SECURITY: Validate account identifier to prevent SSRF attacks
    const accountValidation = validateSnowflakeAccount(account);
    if (!accountValidation.valid) {
      return res.status(400).json({
        success: false,
        message: accountValidation.error
      });
    }

    if (!client_id) {
      return res.status(400).json({
        success: false,
        message: 'Missing client ID'
      });
    }

    const tokenUrl = `https://${account}.snowflakecomputing.com/oauth/token-request`;

    const tokenParams = new URLSearchParams({
      grant_type: 'authorization_code',
      code: code,
      redirect_uri: redirect_uri,
      code_verifier: code_verifier,
      client_id: client_id
    });

    // For LOCAL_APPLICATION, add client_secret per Snowflake docs
    if (client_id === 'LOCAL_APPLICATION') {
      tokenParams.append('client_secret', 'LOCAL_APPLICATION');
    }
    // For custom OAuth clients, add client_secret only if configured in environment
    // This supports both confidential clients (with secret) and public clients (PKCE-only)
    else if (process.env.SNOWFLAKE_OAUTH_CLIENT_SECRET) {
      tokenParams.append('client_secret', process.env.SNOWFLAKE_OAUTH_CLIENT_SECRET);
    }
    // If no client_secret is provided, this is a public client using PKCE-only flow
    // which is secure for client-side applications

    const fetch = (await import('node-fetch')).default;
    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
      },
      body: tokenParams.toString()
    });

    const responseText = await response.text();
    let tokenData;

    try {
      tokenData = JSON.parse(responseText);
    } catch (parseError) {
      return res.status(500).json({
        success: false,
        message: 'Invalid response from Snowflake token endpoint'
      });
    }

    if (!response.ok) {
      // SECURITY: Sanitize error - log server-side, return safe message to client
      const sanitized = sanitizeError({
        message: tokenData.error_description || tokenData.error || 'Token exchange failed',
        details: tokenData,
        statusCode: response.status
      }, 'token_exchange');

      return res.status(response.status).json({
        success: false,
        message: sanitized.message,
        errorId: sanitized.errorId
      });
    }

    if (!tokenData.access_token) {
      // SECURITY: Sanitize error - log full response but return generic message
      const sanitized = sanitizeError({
        message: 'No access token in response',
        response: tokenData
      }, 'token_exchange');

      return res.status(500).json({
        success: false,
        message: sanitized.message,
        errorId: sanitized.errorId
      });
    }

    res.json({
      success: true,
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_in: tokenData.expires_in || 3600,
      token_type: tokenData.token_type || 'Bearer',
      scope: tokenData.scope
    });

  } catch (error) {
    // SECURITY: Sanitize error - never expose internal error messages
    const sanitized = sanitizeError(error, 'token_exchange');

    res.status(500).json({
      success: false,
      message: sanitized.message,
      errorId: sanitized.errorId
    });
  }
});

router.post('/refresh', async (req, res) => {
  try {
    const { refresh_token, account, client_id } = req.body;

    if (!refresh_token) {
      return res.status(400).json({
        success: false,
        message: 'Missing refresh token'
      });
    }

    if (!account) {
      return res.status(400).json({
        success: false,
        message: 'Missing Snowflake account'
      });
    }

    // SECURITY: Validate account identifier to prevent SSRF attacks
    const accountValidation = validateSnowflakeAccount(account);
    if (!accountValidation.valid) {
      return res.status(400).json({
        success: false,
        message: accountValidation.error
      });
    }

    // Use provided client_id or fall back to environment variables
    const finalClientId = client_id || process.env.SNOWFLAKE_OAUTH_CLIENT_ID;
    // IMPORTANT: client_secret should ONLY come from secure environment variables
    // Never accept it from the client request for security reasons
    const finalClientSecret = process.env.SNOWFLAKE_OAUTH_CLIENT_SECRET;

    if (!finalClientId) {
      return res.status(400).json({
        success: false,
        message: 'Missing client ID'
      });
    }

    const tokenUrl = `https://${account}.snowflakecomputing.com/oauth/token-request`;

    const tokenParams = new URLSearchParams({
      grant_type: 'refresh_token',
      refresh_token: refresh_token,
      client_id: finalClientId
    });

    // For LOCAL_APPLICATION, always add client_secret per Snowflake docs
    if (finalClientId === 'LOCAL_APPLICATION') {
      tokenParams.append('client_secret', 'LOCAL_APPLICATION');
    }
    // For custom OAuth clients, add client_secret only if configured
    // This supports both confidential clients (with secret) and public clients (PKCE-only)
    else if (finalClientSecret) {
      tokenParams.append('client_secret', finalClientSecret);
    }
    // If no client_secret, this is a public client refresh token flow

    const fetch = (await import('node-fetch')).default;
    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
      },
      body: tokenParams.toString()
    });

    const responseText = await response.text();
    let tokenData;

    try {
      tokenData = JSON.parse(responseText);
    } catch (parseError) {
      return res.status(500).json({
        success: false,
        message: 'Invalid response from Snowflake token endpoint'
      });
    }

    if (!response.ok) {
      // SECURITY: Sanitize error - log server-side, return safe message to client
      const sanitized = sanitizeError({
        message: tokenData.error_description || tokenData.error || 'Token refresh failed',
        details: tokenData,
        statusCode: response.status
      }, 'token_refresh');

      return res.status(response.status).json({
        success: false,
        message: sanitized.message,
        errorId: sanitized.errorId
      });
    }

    if (!tokenData.access_token) {
      // SECURITY: Sanitize error - log full response but return generic message
      const sanitized = sanitizeError({
        message: 'No access token in response',
        response: tokenData
      }, 'token_refresh');

      return res.status(500).json({
        success: false,
        message: sanitized.message,
        errorId: sanitized.errorId
      });
    }

    res.json({
      success: true,
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token || refresh_token,
      expires_in: tokenData.expires_in || 3600,
      token_type: tokenData.token_type || 'Bearer',
      scope: tokenData.scope
    });

  } catch (error) {
    // SECURITY: Sanitize error - never expose internal error messages
    const sanitized = sanitizeError(error, 'token_refresh');

    res.status(500).json({
      success: false,
      message: sanitized.message,
      errorId: sanitized.errorId
    });
  }
});

router.post('/revoke', async (req, res) => {
  try {
    const { access_token, account } = req.body;

    if (!access_token) {
      return res.status(400).json({
        success: false,
        message: 'Missing access token'
      });
    }

    if (!account) {
      return res.status(400).json({
        success: false,
        message: 'Missing Snowflake account'
      });
    }

    // SECURITY: Validate account identifier to prevent SSRF attacks
    const accountValidation = validateSnowflakeAccount(account);
    if (!accountValidation.valid) {
      return res.status(400).json({
        success: false,
        message: accountValidation.error
      });
    }

    const clientId = process.env.SNOWFLAKE_OAUTH_CLIENT_ID;
    const clientSecret = process.env.SNOWFLAKE_OAUTH_CLIENT_SECRET;

    if (!clientId || !clientSecret) {
      return res.status(500).json({
        success: false,
        message: 'OAuth credentials not configured on server'
      });
    }

    const revokeUrl = `https://${account}.snowflakecomputing.com/oauth/revoke-token`;

    const revokeParams = new URLSearchParams({
      token: access_token,
      client_id: clientId,
      client_secret: clientSecret
    });

    const fetch = (await import('node-fetch')).default;
    await fetch(revokeUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
      },
      body: revokeParams.toString()
    });

    res.json({
      success: true,
      message: 'Token revoked successfully'
    });

  } catch (error) {
    // SECURITY: Sanitize error - log server-side but return generic message
    // Note: We return success=true because logout should proceed even if revocation fails
    const sanitized = sanitizeError(error, 'token_revoke');

    res.json({
      success: true,
      message: 'Logout completed (revocation may have failed)',
      errorId: sanitized.errorId
    });
  }
});

module.exports = router;
