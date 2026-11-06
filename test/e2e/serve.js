'use strict';
// Serves the real webview HTML (same CSP) with a stub of the VS Code webview API.
// Usage: node test/e2e/serve.js <deck.pptx> <dark|light> [maxMB]
// Prints the port on the first line.
const http = require('http');
const fs = require('fs');
const path = require('path');
const { buildHtml } = require('../../src/html');
const { parseDeck } = require('../../src/pptx');
const { readConfig } = require('../../src/protocol');

const ROOT = path.resolve(__dirname, '../..');
const [deckPath, theme = 'dark', maxMB = '150'] = process.argv.slice(2);
const NONCE = 'testnonce123';
const bytes = fs.readFileSync(deckPath);
const fileName = path.basename(deckPath);

const THEMES = {
  dark: { '--vscode-foreground': '#cccccc', '--vscode-editor-background': '#1e1e1e', '--vscode-editorWidget-background': '#252526', '--vscode-sideBar-background': '#252526', '--vscode-panel-border': '#3c3c3c', '--vscode-button-background': '#0e639c', '--vscode-button-foreground': '#ffffff', '--vscode-button-secondaryBackground': '#3a3d41', '--vscode-button-secondaryForeground': '#cccccc', '--vscode-button-secondaryHoverBackground': '#45494e', '--vscode-input-background': '#3c3c3c', '--vscode-input-foreground': '#cccccc', '--vscode-input-border': '#3c3c3c', '--vscode-list-activeSelectionBackground': '#04395e', '--vscode-list-activeSelectionForeground': '#ffffff', '--vscode-list-hoverBackground': '#2a2d2e', '--vscode-editorPane-background': '#1e1e1e', '--vscode-textBlockQuote-background': '#2b2b2b', '--vscode-editor-findMatchHighlightBackground': 'rgba(234,92,0,.45)', '--vscode-font-family': '-apple-system, "Segoe UI", sans-serif', '--vscode-font-size': '13px' },
  light: { '--vscode-foreground': '#333333', '--vscode-editor-background': '#ffffff', '--vscode-font-family': '-apple-system, "Segoe UI", sans-serif', '--vscode-font-size': '13px' },
};

function stubJs() {
  const deckOrError = (() => { try { return { deck: parseDeck(bytes) }; } catch (e) { return { error: e.message }; } })();
  const config = readConfig((k, d) => (k === 'maxFileSizeMB' ? Number(maxMB) : d));
  return `
window.__sent = [];
(function(){
  const vars = ${JSON.stringify(THEMES[theme] || THEMES.dark)};
  for (const [k, v] of Object.entries(vars)) document.documentElement.style.setProperty(k, v);
  const b64 = ${JSON.stringify(bytes.toString('base64'))};
  const payload = ${JSON.stringify(deckOrError)};
  const config = ${JSON.stringify(config)};
  const size = ${bytes.length};
  let force = false;
  const post = (m) => window.postMessage(m, '*');
  function load() {
    if (size > config.maxFileSizeBytes && !force) { post({ type: 'tooLarge', size, limit: config.maxFileSizeBytes }); return; }
    if (payload.error) { post({ type: 'error', message: payload.error }); return; }
    post({ type: 'deck', deck: payload.deck, fileName: ${JSON.stringify(fileName)} });
    const bin = atob(b64); const u8 = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    if (config.renderSlides) post({ type: 'bytes', data: u8 });
  }
  let state = null;
  window.acquireVsCodeApi = () => ({
    getState: () => state, setState: (s) => { state = s; },
    postMessage: (m) => {
      window.__sent.push(m);
      if (m.type === 'ready') { post({ type: 'init', config, fileName: ${JSON.stringify(fileName)} }); setTimeout(load, 0); }
      if (m.type === 'loadAnyway') { force = true; load(); }
    },
  });
})();`;
}

const server = http.createServer((req, res) => {
  const url = decodeURIComponent(req.url.split('?')[0]);
  if (url === '/harness.html') {
    const base = `http://127.0.0.1:${server.address().port}`;
    let html = buildHtml({ css: base + '/media/viewer.css', core: base + '/media/core.js', viewer: base + '/media/viewer.js', preview: base + '/media/vendor/pptx-preview.umd.js' }, base, NONCE);
    html = html.replace(`<script nonce="${NONCE}"`, `<script nonce="${NONCE}" src="${base}/stub.js"></script>\n<script nonce="${NONCE}"`);
    res.writeHead(200, { 'content-type': 'text/html' });
    return res.end(html);
  }
  if (url === '/stub.js') { res.writeHead(200, { 'content-type': 'text/javascript' }); return res.end(stubJs()); }
  if (url.startsWith('/media/')) {
    const f = path.join(ROOT, path.normalize(url).replace(/^(\.\.[/\\])+/, ''));
    if (f.startsWith(path.join(ROOT, 'media')) && fs.existsSync(f)) {
      const type = f.endsWith('.css') ? 'text/css' : 'text/javascript';
      res.writeHead(200, { 'content-type': type });
      return res.end(fs.readFileSync(f));
    }
  }
  res.writeHead(404);
  res.end();
});
server.listen(0, '127.0.0.1', () => console.log(server.address().port));
