const express = require('express');
const router = express.Router();

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
      return res.status(response.status).json({
        success: false,
        message: tokenData.error_description || tokenData.error || 'Token exchange failed',
        error: tokenData.error,
        details: tokenData
      });
    }

    if (!tokenData.access_token) {
      return res.status(500).json({
        success: false,
        message: 'No access token received from Snowflake'
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
    res.status(500).json({
      success: false,
      message: 'Internal server error during token exchange',
      error: error.message
    });
  }
});

router.post('/refresh', async (req, res) => {
  try {
    const { refresh_token, account, client_id, client_secret } = req.body;

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

    // Use provided client_id or fall back to environment variables
    const finalClientId = client_id || process.env.SNOWFLAKE_OAUTH_CLIENT_ID;
    const finalClientSecret = client_secret || process.env.SNOWFLAKE_OAUTH_CLIENT_SECRET;

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
    } else if (finalClientSecret) {
      tokenParams.append('client_secret', finalClientSecret);
    }

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
      return res.status(response.status).json({
        success: false,
        message: tokenData.error_description || tokenData.error || 'Token refresh failed',
        error: tokenData.error
      });
    }

    if (!tokenData.access_token) {
      return res.status(500).json({
        success: false,
        message: 'No access token received from Snowflake'
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
    res.status(500).json({
      success: false,
      message: 'Internal server error during token refresh',
      error: error.message
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
    res.json({
      success: true,
      message: 'Logout completed (revocation may have failed)'
    });
  }
});

module.exports = router;
