// Minimal local dev server: serves public/ and proxies /.netlify/functions/*
// to the same handler code Netlify runs in production.
// Reads credentials from .env automatically — same pattern as autotask-report.

const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  fs.readFileSync(envPath, 'utf8').split('\n').forEach(line => {
    const [k, ...v] = line.split('=');
    if (k && v.length) process.env[k.trim()] = v.join('=').trim();
  });
}

const PORT = process.env.PORT || 8789;
const PUBLIC_DIR = path.join(__dirname, 'public');
const FUNCTIONS_DIR = path.join(__dirname, 'netlify', 'functions');

const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json' };

const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);

  if (parsed.pathname.startsWith('/.netlify/functions/')) {
    const fnName = parsed.pathname.replace('/.netlify/functions/', '');
    const fnPath = path.join(FUNCTIONS_DIR, `${fnName}.js`);
    if (!fs.existsSync(fnPath)) { res.writeHead(404); return res.end('Function not found'); }
    delete require.cache[require.resolve(fnPath)];
    const fn = require(fnPath);
    try {
      const result = await fn.handler({ queryStringParameters: parsed.query, httpMethod: req.method, headers: req.headers });
      res.writeHead(result.statusCode || 200, { 'Content-Type': 'application/json', ...(result.headers || {}) });
      res.end(result.body);
    } catch (err) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  let filePath = path.join(PUBLIC_DIR, parsed.pathname === '/' ? 'index.html' : parsed.pathname);
  if (!fs.existsSync(filePath)) { res.writeHead(404); return res.end('Not found'); }
  const ext = path.extname(filePath);
  res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
  fs.createReadStream(filePath).pipe(res);
});

server.listen(PORT, () => console.log(`jarvis-dashboard running at http://localhost:${PORT}`));
