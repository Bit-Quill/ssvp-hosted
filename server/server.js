/**
 * Express Server for Snowflake Excel Plugin
 * Handles OAuth callbacks and API proxying
 */

const express = require('express');
const path = require('path');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// ========== Middleware ==========

// CORS configuration
app.use(cors({
  origin: process.env.ALLOWED_ORIGINS ? process.env.ALLOWED_ORIGINS.split(',') : '*',
  credentials: true
}));

// Body parsing
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Security headers
app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'SAMEORIGIN');
  res.setHeader('X-XSS-Protection', '1; mode=block');
  res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
  next();
});

// ========== Routes ==========

// OAuth authentication routes
const authRoutes = require('./routes/auth');
app.use('/api/auth', authRoutes);

// Cortex Analyst proxy routes
const cortexRoutes = require('./routes/cortex');
app.use('/api/cortex', cortexRoutes);

// ========== OAuth Callback Handler ==========

/**
 * OAuth callback endpoint for Snowflake OAuth
 *
 * When Snowflake redirects to https://localhost:3000/auth/callback?code=...
 * this endpoint receives the authorization code and displays it to the user.
 *
 * The Excel add-in's popup window will extract the code and handle token exchange.
 */
app.get('/auth/callback', (req, res) => {
  const { code, state, error, error_description } = req.query;

  // Send the callback HTML page
  // This page will use window.opener.postMessage to send code to the add-in
  res.sendFile(path.join(__dirname, '../src/auth/callback.html'));
});

// Serve static files (task pane, assets, etc.)
if (process.env.NODE_ENV === 'production') {
  app.use(express.static(path.join(__dirname, '../dist')));
} else {
  app.use(express.static(path.join(__dirname, '../src')));
}

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV || 'development'
  });
});

// API status endpoint
app.get('/api/status', (req, res) => {
  const isConfigured = !!(
    process.env.SNOWFLAKE_OAUTH_CLIENT_ID &&
    process.env.SNOWFLAKE_OAUTH_CLIENT_SECRET
  );

  res.json({
    status: 'online',
    oauth_configured: isConfigured,
    version: '1.0.0'
  });
});

// Catch-all route for SPA (serves taskpane.html for all unmatched routes)
app.get('*', (req, res) => {
  if (process.env.NODE_ENV === 'production') {
    res.sendFile(path.join(__dirname, '../dist/taskpane.html'));
  } else {
    res.sendFile(path.join(__dirname, '../src/taskpane/taskpane.html'));
  }
});

// ========== Error Handling ==========

// 404 handler
app.use((req, res) => {
  res.status(404).json({
    success: false,
    message: 'Route not found'
  });
});

// Global error handler
app.use((err, req, res, next) => {
  res.status(err.status || 500).json({
    success: false,
    message: err.message || 'Internal server error',
    ...(process.env.NODE_ENV !== 'production' && { stack: err.stack })
  });
});

// ========== Server Startup ==========

// Start server (HTTP for development, HTTPS in production via reverse proxy)
app.listen(PORT, () => {
  // Server started
});

// Graceful shutdown
process.on('SIGTERM', () => {
  process.exit(0);
});

process.on('SIGINT', () => {
  process.exit(0);
});

module.exports = app;
