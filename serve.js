const http = require('http');
const fs = require('fs');
const path = require('path');
const build = require('./build');

const PORT = 3000;
const ROOT = __dirname;

const MIME = {
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'application/javascript',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.json': 'application/json',
};

// SSE clients waiting for reload
const clients = [];

// Debounce rebuild to avoid double-fires from fs.watch
let debounceTimer = null;
const watchDirs = ['slides', 'css', 'js'];
watchDirs.forEach(dir => {
  fs.watch(path.join(ROOT, dir), { recursive: true }, (event, filename) => {
    if (!filename) return;
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      console.log(`[watch] ${dir}/${filename} changed — rebuilding...`);
      try {
        build();
        console.log(`[reload] notifying ${clients.length} client(s)...`);
        clients.forEach(res => res.write('data: reload\n\n'));
      } catch (err) {
        console.error('[build error]', err.message);
      }
    }, 200);
  });
});

// Injected into served HTML — listens for SSE reload event
const RELOAD_SCRIPT = `<script>
new EventSource('/__livereload').onmessage = function() {
  sessionStorage.setItem('__slide', typeof current !== 'undefined' ? current : 1);
  location.reload();
};
window.addEventListener('DOMContentLoaded', function() {
  var s = sessionStorage.getItem('__slide');
  if (s) { sessionStorage.removeItem('__slide'); goToSlide(Number(s)); }
});
</script>`;

const server = http.createServer((req, res) => {
  // SSE endpoint for live reload
  if (req.url === '/__livereload') {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
    });
    res.write('\n');
    clients.push(res);
    req.on('close', () => {
      const i = clients.indexOf(res);
      if (i !== -1) clients.splice(i, 1);
    });
    return;
  }

  let filePath = req.url === '/' ? '/index.html' : req.url;
  filePath = path.join(ROOT, filePath);
  const ext = path.extname(filePath);
  const contentType = MIME[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    if (ext === '.html') {
      const html = data.toString().replace('</body>', RELOAD_SCRIPT + '\n</body>');
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(html);
      return;
    }
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`\n  Serving at http://localhost:${PORT}`);
  console.log(`  Watching: ${watchDirs.join(', ')}`);
  console.log(`  Live reload: enabled (SSE)\n`);
});
