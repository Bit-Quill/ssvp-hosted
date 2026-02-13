// API Configuration
// This file is configured at build time by webpack

const config = {
  // API Base URL - configured via webpack DefinePlugin
  API_BASE_URL: process.env.API_BASE_URL || 'http://localhost:3000',

  // OAuth Redirect URI - configured via webpack DefinePlugin
  OAUTH_REDIRECT_URI: process.env.OAUTH_REDIRECT_URI || 'http://localhost:3000/auth/callback',

  // Environment
  NODE_ENV: process.env.NODE_ENV || 'development',

  // Helper functions
  getApiUrl(path) {
    // Ensure path starts with /
    const cleanPath = path.startsWith('/') ? path : `/${path}`;
    return `${this.API_BASE_URL}${cleanPath}`;
  },

  isProduction() {
    return this.NODE_ENV === 'production';
  },

  isDevelopment() {
    return this.NODE_ENV === 'development';
  }
};

export default config;
