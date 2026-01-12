const express = require('express');
const path = require('path');
const cors = require('cors');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 80;

// SECURITY: CORS configuration with strict origin validation
// IMPORTANT: Never use wildcard '*' with credentials: true
// This configuration requires ALLOWED_ORIGINS to be set in production
const allowedOrigins = process.env.ALLOWED_ORIGINS
  ? process.env.ALLOWED_ORIGINS.split(',')
  : process.env.NODE_ENV === 'production'
    ? [] // Production MUST set ALLOWED_ORIGINS explicitly - fail secure
    : [
        // Development defaults - safe for local development only
        'https://localhost:3000',
        'http://localhost:3000',
        'https://127.0.0.1:3000',
        'http://127.0.0.1:3000',
        'https://localhost',
        'http://localhost'
      ];

// Validate that production has ALLOWED_ORIGINS configured
if (process.env.NODE_ENV === 'production' && allowedOrigins.length === 0) {
  console.error('ERROR: ALLOWED_ORIGINS environment variable must be set in production mode');
  console.error('Example: ALLOWED_ORIGINS=https://yourdomain.com,https://app.yourdomain.com');
  process.exit(1);
}

app.use(cors({
  origin: (origin, callback) => {
    // Allow requests with no origin (like mobile apps, curl, Postman)
    if (!origin) {
      return callback(null, true);
    }

    // Check if origin is in allowlist
    if (allowedOrigins.includes(origin)) {
      callback(null, true);
    } else {
      console.warn(`CORS: Blocked request from unauthorized origin: ${origin}`);
      callback(new Error(`CORS policy: Origin ${origin} is not allowed`));
    }
  },
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
