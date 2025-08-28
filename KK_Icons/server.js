/* eslint-disable no-undef */
const express = require('express');
const path = require('path');
const { createProxyMiddleware } = require('http-proxy-middleware');
require('dotenv').config({ path: path.resolve(__dirname, '.env') });

const app = express();
const port = process.env.PORT || 8080;
const apiTarget = 'https://public-api.streamlinehq.com/v1';

// Log presence (not value) of env at startup
if (process.env.STREAMLINE_API_KEY) {
  console.log('[server] STREAMLINE_API_KEY detected (length:', String(process.env.STREAMLINE_API_KEY.length), ')');
} else {
  console.warn('[server] STREAMLINE_API_KEY is NOT set at process start');
}

// Serve static files from dist
app.use(express.static(path.join(__dirname, 'dist')));

// Proxy for Streamline API, inject API key header
app.use('/api/streamline', createProxyMiddleware({
  target: apiTarget,
  changeOrigin: true,
  pathRewrite: { '^/api/streamline': '' },
  secure: true,
  logLevel: 'debug',
  onProxyReq: (proxyReq, req, res) => {
    const apiKey = process.env.STREAMLINE_API_KEY || '';
    if (!apiKey) {
      console.warn('[server] STREAMLINE_API_KEY not set');
    }
    proxyReq.setHeader('x-api-key', apiKey);
    console.log('[server] Injecting x-api-key for', req.method, req.url, 'keyLen=', apiKey ? String(apiKey.length) : '0');
    if (req.url.includes('/download/svg')) {
      proxyReq.setHeader('accept', 'image/svg+xml');
    } else {
      proxyReq.setHeader('accept', 'application/json');
    }
  }
}));

// Fallback to taskpane.html for root
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'taskpane.html'));
});

// Health check
app.get('/healthz', (req, res) => res.send('OK'));

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});


