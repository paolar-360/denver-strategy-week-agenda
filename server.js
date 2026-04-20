// Local dev server (no Vercel login required).
// Serves /public as static files and emulates /api/* serverless functions.
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = __dirname;
const API_DIR = path.join(__dirname, 'api');
const STATIC_ALLOW = new Set(['.html', '.css', '.js', '.json', '.svg', '.png', '.jpg', '.jpeg', '.ico', '.webp', '.woff', '.woff2']);

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

function sendStatic(res, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const type = MIME[ext] || 'application/octet-stream';
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': type, 'Cache-Control': 'no-store' });
    res.end(data);
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  let pathname = decodeURIComponent(url.pathname);

  // API route → load function from api/
  if (pathname.startsWith('/api/')) {
    const name = pathname.replace(/^\/api\//, '').replace(/\/+$/, '');
    const fnPath = path.join(API_DIR, `${name}.js`);
    if (!fs.existsSync(fnPath)) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('API not found');
      return;
    }
    try {
      // Bust require cache so edits to api/* are reflected on reload
      delete require.cache[require.resolve(fnPath)];
      const handler = require(fnPath);
      const fn = typeof handler === 'function' ? handler : handler.default;
      await fn(req, res);
    } catch (err) {
      console.error(err);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  if (pathname === '/') pathname = '/index.html';

  // Only allow known static asset types so we don't expose server.js, package.json, etc.
  const ext = path.extname(pathname).toLowerCase();
  if (!STATIC_ALLOW.has(ext)) {
    sendStatic(res, path.join(PUBLIC_DIR, 'index.html'));
    return;
  }

  const filePath = path.join(PUBLIC_DIR, pathname);
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403); res.end('Forbidden'); return;
  }
  fs.stat(filePath, (err, stat) => {
    if (err || !stat.isFile()) {
      sendStatic(res, path.join(PUBLIC_DIR, 'index.html'));
    } else {
      sendStatic(res, filePath);
    }
  });
});

server.listen(PORT, () => {
  console.log(`\n→ Local server: http://localhost:${PORT}\n`);
});
