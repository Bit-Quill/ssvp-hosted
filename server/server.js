const express = require('express');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

// Load dotenv if available (optional - only needed for development)
try {
  require('dotenv').config();
} catch (e) {
  // dotenv not available - this is fine for production executables
}

const app = express();
const PORT = process.env.PORT || 80;

// Detect if running as standalone executable or normal Node.js
const isExecutable = typeof process.pkg !== 'undefined';
const baseDir = isExecutable
  ? path.dirname(process.execPath) // For executables in root: use executable's directory
  : path.join(__dirname, '..'); // For normal node: use parent directory of server/

console.log('Running mode:', isExecutable ? 'Standalone Executable' : 'Node.js');
console.log('Base directory:', baseDir);

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
        'https://localhost:80',
        'http://localhost:80',
        'https://127.0.0.1:80',
        'http://127.0.0.1:80',
        'https://localhost',
        'http://localhost',
        'https://127.0.0.1',
        'http://127.0.0.1'
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
    // Prefer dist folder if it exists (for distribution packages)
    const distCallback = path.join(baseDir, 'dist/auth/callback.html');
    const srcCallback = path.join(baseDir, 'src/auth/callback.html');

    if (fs.existsSync(distCallback)) {
      return res.sendFile(distCallback);
    } else if (fs.existsSync(srcCallback)) {
      return res.sendFile(srcCallback);
    } else {
      return res.status(500).send('ERROR: callback.html not found');
    }
  }

  next();
};

app.get('/', handleOAuthCallback);
app.get('/auth/callback', handleOAuthCallback);

// Serve static files - prefer dist folder if it exists (for distribution packages)
const distPath = path.join(baseDir, 'dist');
const srcPath = path.join(baseDir, 'src');

if (fs.existsSync(distPath)) {
  console.log('Serving files from:', distPath);
  app.use(express.static(distPath));
} else if (fs.existsSync(srcPath)) {
  console.log('Serving files from:', srcPath);
  app.use(express.static(srcPath));
} else {
  console.error('ERROR: Neither dist nor src folder found');
  console.error('Checked paths:');
  console.error('  -', distPath);
  console.error('  -', srcPath);
  process.exit(1);
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
  // Serve taskpane.html - prefer dist folder if it exists
  const distHtml = path.join(baseDir, 'dist/taskpane.html');
  const srcHtml = path.join(baseDir, 'src/taskpane/taskpane.html');

  if (fs.existsSync(distHtml)) {
    res.sendFile(distHtml);
  } else if (fs.existsSync(srcHtml)) {
    res.sendFile(srcHtml);
  } else {
    res.status(500).send('ERROR: taskpane.html not found');
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

const server = app.listen(PORT, () => {
  console.log('========================================');
  console.log('Snowflake Excel Server');
  console.log('========================================');
  console.log(`Server running on port ${PORT}`);
  console.log(`Status: http://localhost:${PORT}/health`);
  console.log('');
  console.log('✓ Server is ready!');
  console.log('✓ You can now use the Excel add-in');
  console.log('');
  console.log('Keep this window open while using the plugin.');
  console.log('Press Ctrl+C to stop the server.');
  console.log('========================================');
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error('');
    console.error('========================================');
    console.error('ERROR: Port 80 is already in use!');
    console.error('========================================');
    console.error('');
    console.error('Another application is using port 80.');
    console.error('');
    console.error('Common causes:');
    console.error('  - Another instance of this server is running');
    console.error('  - nginx web server');
    console.error('  - Apache web server');
    console.error('  - IIS (Windows)');
    console.error('  - Skype (Windows)');
    console.error('');
    console.error('Solutions:');
    console.error('');
    console.error('Mac/Linux - Check what\'s using port 80:');
    console.error('  sudo lsof -i :80');
    console.error('');
    console.error('Mac - Stop nginx:');
    console.error('  sudo nginx -s stop');
    console.error('');
    console.error('Mac - Stop Apache:');
    console.error('  sudo apachectl stop');
    console.error('');
    console.error('Windows - Check what\'s using port 80:');
    console.error('  netstat -ano | findstr :80');
    console.error('');
    console.error('Windows - Stop IIS:');
    console.error('  iisreset /stop');
    console.error('');
    console.error('For more help, see TROUBLESHOOTING.md');
    console.error('========================================');
    console.error('');
    process.exit(1);
  } else if (err.code === 'EACCES') {
    console.error('');
    console.error('========================================');
    console.error('ERROR: Permission denied for port 80');
    console.error('========================================');
    console.error('');
    console.error('Port 80 requires administrator/root privileges.');
    console.error('');
    console.error('Solutions:');
    console.error('');
    console.error('Mac/Linux - Run with sudo:');
    console.error('  sudo ./snowflake-excel-server-macos');
    console.error('');
    console.error('Windows - Run as Administrator:');
    console.error('  Right-click the executable → "Run as administrator"');
    console.error('');
    console.error('========================================');
    console.error('');
    process.exit(1);
  } else {
    console.error('');
    console.error('========================================');
    console.error('ERROR: Failed to start server');
    console.error('========================================');
    console.error('');
    console.error('Error:', err.message);
    console.error('');
    console.error('For help, see TROUBLESHOOTING.md');
    console.error('========================================');
    console.error('');
    process.exit(1);
  }
});

process.on('SIGTERM', () => {
  console.log('\nShutting down server...');
  server.close(() => {
    process.exit(0);
  });
});

process.on('SIGINT', () => {
  console.log('\nShutting down server...');
  server.close(() => {
    process.exit(0);
  });
});

module.exports = app;
