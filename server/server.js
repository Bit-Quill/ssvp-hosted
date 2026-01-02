const express = require('express');
const path = require('path');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 80;

app.use(cors({
  origin: process.env.ALLOWED_ORIGINS ? process.env.ALLOWED_ORIGINS.split(',') : '*',
  credentials: true
}));

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'SAMEORIGIN');
  res.setHeader('X-XSS-Protection', '1; mode=block');
  res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
  next();
});

const authRoutes = require('./routes/auth');
app.use('/api/auth', authRoutes);

const handleOAuthCallback = (req, res, next) => {
  const { code, error } = req.query;

  if (code || error) {
    return res.sendFile(path.join(__dirname, '../src/auth/callback.html'));
  }

  next();
};

app.get('/', handleOAuthCallback);
app.get('/auth/callback', handleOAuthCallback);

if (process.env.NODE_ENV === 'production') {
  app.use(express.static(path.join(__dirname, '../dist')));
} else {
  app.use(express.static(path.join(__dirname, '../src')));
}

app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV || 'development'
  });
});

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

app.get('*', (req, res) => {
  if (process.env.NODE_ENV === 'production') {
    res.sendFile(path.join(__dirname, '../dist/taskpane.html'));
  } else {
    res.sendFile(path.join(__dirname, '../src/taskpane/taskpane.html'));
  }
});

app.use((req, res) => {
  res.status(404).json({
    success: false,
    message: 'Route not found'
  });
});

app.use((err, req, res, next) => {
  res.status(err.status || 500).json({
    success: false,
    message: err.message || 'Internal server error',
    ...(process.env.NODE_ENV !== 'production' && { stack: err.stack })
  });
});

app.listen(PORT, () => {});

process.on('SIGTERM', () => {
  process.exit(0);
});

process.on('SIGINT', () => {
  process.exit(0);
});

module.exports = app;
