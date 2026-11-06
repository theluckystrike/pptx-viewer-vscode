'use strict';
// Drives the real custom editor provider with a stub of the vscode API: ready -> init, deck, bytes;
// copy as Markdown -> clipboard; size limit -> tooLarge, then loadAnyway.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const Module = require('module');

const FIX = path.join(__dirname, 'fixtures');
const settings = {};
let clipboard = '';
const commands = new Map();
const uri = (p) => ({ path: p, fsPath: p, toString: () => 'file://' + p });
const vscodeStub = {
  Uri: { joinPath: (u, ...p) => uri(path.posix.join(u.path, ...p)), parse: (s) => uri(s) },
  RelativePattern: class { constructor(b, p) { this.base = b; this.pattern = p; } },
  workspace: {
    getConfiguration: () => ({ get: (k, d) => (k in settings ? settings[k] : d) }),
    fs: { stat: async (u) => ({ size: fs.statSync(u.path).size }), readFile: async (u) => fs.readFileSync(u.path) },
    createFileSystemWatcher: () => ({ onDidChange: () => ({ dispose() {} }), onDidCreate: () => ({ dispose() {} }), dispose() {} }),
    onDidChangeConfiguration: () => ({ dispose() {} }),
  },
  window: {
    registerCustomEditorProvider: () => ({ dispose() {} }),
    setStatusBarMessage: () => ({ dispose() {} }),
    showOpenDialog: async () => undefined,
  },
  env: { clipboard: { writeText: async (t) => { clipboard = t; } } },
  commands: { registerCommand: (id, fn) => { commands.set(id, fn); return { dispose() {} }; }, executeCommand: async () => {} },
};
const origLoad = Module._load;
Module._load = function (req, ...rest) { return req === 'vscode' ? vscodeStub : origLoad.call(this, req, ...rest); };
const ext = require('../src/extension');

function openPanel(provider, file) {
  const sent = [];
  let onMsg = null;
  const panel = {
    active: true, visible: true,
    webview: {
      options: null, html: '', cspSource: 'vscode-webview://test',
      asWebviewUri: (u) => ({ toString: () => 'vscode-webview://test' + u.path }),
      postMessage: async (m) => { sent.push(m); return true; },
      onDidReceiveMessage: (fn) => { onMsg = fn; return { dispose() {} }; },
    },
    onDidChangeViewState: () => ({ dispose() {} }),
    onDidDispose: () => ({ dispose() {} }),
  };
  const doc = provider.openCustomDocument(uri(path.join(FIX, file)));
  provider.resolveCustomEditor(doc, panel);
  return { panel, sent, post: (m) => onMsg(m) };
}

const wait = () => new Promise((r) => setTimeout(r, 30));

test('activate registers the commands', () => {
  const { provider } = ext.activate({ subscriptions: [], extensionUri: uri('/ext') });
  for (const id of ['pptxViewer.copySlideMarkdown', 'pptxViewer.copyDeckMarkdown', 'pptxViewer.find', 'pptxViewer.toggleNotes', 'pptxViewer.openWith']) assert.ok(commands.has(id), id);
  assert.ok(provider);
});

test('ready -> init, deck (text first), then bytes; webview options are locked down', async () => {
  const provider = new ext.PptxEditorProvider({ extensionUri: uri('/ext') });
  const p = openPanel(provider, 'quarterly-review.pptx');
  assert.deepEqual(p.panel.webview.options.localResourceRoots.map((u) => u.path), ['/ext/media']);
  assert.match(p.panel.webview.html, /connect-src 'none'/);
  p.post({ type: 'ready' });
  await wait();
  assert.deepEqual(p.sent.map((m) => m.type), ['init', 'deck', 'bytes']);
  assert.equal(p.sent[1].deck.slides.length, 6);
  assert.equal(Object.getPrototypeOf(p.sent[2].data), Uint8Array.prototype);
});

test('copySlideMarkdown and copyDeckMarkdown write the clipboard; bad messages are ignored', async () => {
  const provider = new ext.PptxEditorProvider({ extensionUri: uri('/ext') });
  const p = openPanel(provider, 'quarterly-review.pptx');
  p.post({ type: 'ready' });
  await wait();
  await p.post({ type: 'copySlideMarkdown', slide: 3 });
  assert.ok(clipboard.startsWith('## Latency by region'));
  await p.post({ type: 'copyDeckMarkdown' });
  assert.ok(clipboard.startsWith('# quarterly-review.pptx'));
  clipboard = 'unchanged';
  await p.post({ type: 'copySlideMarkdown', slide: 99 });
  await p.post({ type: 'copySlideMarkdown', slide: '1' });
  await p.post({ type: 'evil', code: 'x' });
  assert.equal(clipboard, 'unchanged');
});

test('size limit sends tooLarge, loadAnyway opens it', async () => {
  settings.maxFileSizeMB = 0.01;
  try {
    const provider = new ext.PptxEditorProvider({ extensionUri: uri('/ext') });
    const p = openPanel(provider, 'quarterly-review.pptx');
    p.post({ type: 'ready' });
    await wait();
    assert.deepEqual(p.sent.map((m) => m.type), ['init', 'tooLarge']);
    p.post({ type: 'loadAnyway' });
    await wait();
    assert.deepEqual(p.sent.map((m) => m.type).slice(2), ['deck', 'bytes']);
  } finally { delete settings.maxFileSizeMB; }
});

test('renderSlides off sends text only; a broken file sends an error', async () => {
  settings.renderSlides = false;
  try {
    const provider = new ext.PptxEditorProvider({ extensionUri: uri('/ext') });
    const p = openPanel(provider, 'quarterly-review.pptx');
    p.post({ type: 'ready' });
    await wait();
    assert.deepEqual(p.sent.map((m) => m.type), ['init', 'deck']);
  } finally { delete settings.renderSlides; }
  const bad = path.join(FIX, '..', 'not-a-deck.tmp.pptx');
  fs.writeFileSync(bad, 'plain text, not a zip');
  try {
    const provider = new ext.PptxEditorProvider({ extensionUri: uri('/ext') });
    const p = openPanel(provider, path.relative(FIX, bad));
    p.post({ type: 'ready' });
    await wait();
    assert.equal(p.sent[1].type, 'error');
    assert.match(p.sent[1].message, /not a ZIP package/);
  } finally { fs.unlinkSync(bad); }
});

// ---------- support links (Sponsor, More free viewers) ----------
test('support: package.json has the sponsor field, the two commands and the setting', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8'));
  const support = require('../src/support');
  assert.equal(pkg.sponsor.url, 'https://github.com/sponsors/theluckystrike');
  assert.equal(pkg.sponsor.url, support.sponsorUrl());
  const ids = pkg.contributes.commands.map((c) => c.command);
  assert.ok(ids.includes('pptxViewer.supportDeveloper') && ids.includes('pptxViewer.moreViewers'));
  const cfg = Array.isArray(pkg.contributes.configuration) ? pkg.contributes.configuration : [pkg.contributes.configuration];
  const props = Object.assign({}, ...cfg.map((c) => c.properties));
  assert.equal(props['pptxViewer.showSupportLinks'].default, true);
  assert.ok(support.otherProducts('zovo.pptx-slide-viewer').every((p) => p.id !== 'zovo.pptx-slide-viewer'));
});

test('support: the footer shows by default and pptxViewer.showSupportLinks=false hides it', () => {
  ext.activate({ subscriptions: [], extensionUri: uri('/ext') });
  assert.ok(commands.has('pptxViewer.supportDeveloper') && commands.has('pptxViewer.moreViewers'));
  for (const [value, shown] of [[undefined, true], [false, false]]) {
    if (value === undefined) delete settings.showSupportLinks; else settings.showSupportLinks = value;
    try {
      const provider = new ext.PptxEditorProvider({ extensionUri: uri('/ext') });
      const p = openPanel(provider, 'quarterly-review.pptx');
      const html = p.panel.webview.html;
      assert.equal(/id="zovo-support"/.test(html), shown, String(value));
      assert.equal(html.includes('command:pptxViewer.moreViewers'), shown);
      assert.deepEqual(p.panel.webview.options.enableCommandUris, ['pptxViewer.supportDeveloper', 'pptxViewer.moreViewers', 'workbench.action.openSettings']);
    } finally { delete settings.showSupportLinks; }
  }
});
