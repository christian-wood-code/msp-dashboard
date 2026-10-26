// Minimal local dev server: serves public/ and proxies /.netlify/functions/*
// to the same handler code Netlify runs in production.
// Reads credentials from .env automatically — same pattern as autotask-report.

const http = require('http');
const fs = require('fs');
const path = require('path');

const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  fs.readFileSync(envPath, 'utf8').split('\n').forEach(line => {
    const [k, ...v] = line.split('=');
    if (k && v.length) process.env[k.trim()] = v.join('=').trim();
  });
}

const PORT = process.env.PORT || 8789;
// Loopback only by default: this server proxies the Netlify functions with live
// credentials and has no auth, so it must not be reachable from the LAN.
// Set HOST=0.0.0.0 only if you deliberately want that.
const HOST = process.env.HOST || '127.0.0.1';
// Reject non-loopback Host headers (blocks DNS-rebinding: a malicious page
// resolving its own hostname to 127.0.0.1 to read our responses).
const LOOPBACK_HOST = /^(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$/i;
const PUBLIC_DIR = path.join(__dirname, 'public');
const FUNCTIONS_DIR = path.join(__dirname, 'netlify', 'functions');

const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json' };

const server = http.createServer(async (req, res) => {
  if (HOST === '127.0.0.1' && !LOOPBACK_HOST.test(req.headers.host || '')) {
    res.writeHead(403); res.end('Forbidden: unexpected Host header'); return;
  }
  let parsed;
  try { parsed = new URL(req.url, 'http://localhost'); } catch { res.writeHead(400); res.end('Bad request'); return; }

  if (parsed.pathname.startsWith('/.netlify/functions/')) {
    const fnName = parsed.pathname.replace('/.netlify/functions/', '');
    if (!/^[A-Za-z0-9_-]+$/.test(fnName)) { res.writeHead(404); return res.end('Function not found'); }
    const fnPath = path.join(FUNCTIONS_DIR, `${fnName}.js`);
    if (!fs.existsSync(fnPath)) { res.writeHead(404); return res.end('Function not found'); }
    delete require.cache[require.resolve(fnPath)];
    const fn = require(fnPath);
    try {
      const result = await fn.handler({ queryStringParameters: Object.fromEntries(parsed.searchParams), httpMethod: req.method, headers: req.headers });
      res.writeHead(result.statusCode || 200, { 'Content-Type': 'application/json', ...(result.headers || {}) });
      res.end(result.body);
    } catch (err) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  let rel;
  try { rel = decodeURIComponent(parsed.pathname); } catch { res.writeHead(400); return res.end('Bad request'); }
  if (rel.includes('\0')) { res.writeHead(400); return res.end('Bad request'); }
  let filePath = path.resolve(PUBLIC_DIR, '.' + (rel === '/' ? '/index.html' : rel));
  // Never serve anything outside public/ ("/../.env" style traversal).
  if (filePath !== PUBLIC_DIR && !filePath.startsWith(PUBLIC_DIR + path.sep)) { res.writeHead(403); return res.end('Forbidden'); }
  // Regular files only (createReadStream on a directory would crash this server).
  if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) { res.writeHead(404); return res.end('Not found'); }
  const ext = path.extname(filePath);
  res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
  fs.createReadStream(filePath).pipe(res);
});

server.listen(PORT, HOST, () => console.log(`jarvis-dashboard running at http://localhost:${PORT}`));
