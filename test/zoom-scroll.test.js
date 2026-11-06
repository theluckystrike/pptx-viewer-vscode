'use strict';
// Zoom and the scroll view (0.2.0): the pure helpers, the viewPrefs message, the per-file memory in
// workspaceState, the commands and keybindings in package.json, and that the webview handles each command.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const Module = require('module');
const C = require('../media/core');
const { parseWebviewMessage, readConfig } = require('../src/protocol');
const { buildHtml } = require('../src/html');

const ROOT = path.join(__dirname, '..');
const FIX = path.join(__dirname, 'fixtures');

// ---------- helpers in media/core.js ----------
test('zoomScale: fit, width and a fixed scale', () => {
  // A 960 x 540 slide in a 480 x 400 box.
  assert.equal(C.zoomScale('fit', 480, 400, 960, 540), 0.5);
  assert.equal(C.zoomScale('width', 480, 100, 960, 540), 0.5);
  assert.equal(C.zoomScale('width', 1920, 100, 960, 540), 2);
  assert.equal(C.zoomScale(1.5, 480, 400, 960, 540), 1.5);
  assert.equal(C.zoomScale('nonsense', 480, 400, 960, 540), 0.5, 'unknown zoom falls back to fit');
  assert.equal(C.zoomScale('width', 0, 0, 960, 540), 1);
});

test('normalizeZoom clamps to 10% .. 400% and keeps fit and width', () => {
  assert.equal(C.normalizeZoom('fit'), 'fit');
  assert.equal(C.normalizeZoom('width'), 'width');
  assert.equal(C.normalizeZoom(9), 4);
  assert.equal(C.normalizeZoom(0.01), 0.1);
  assert.equal(C.normalizeZoom(-1), 'fit');
  assert.equal(C.normalizeZoom(NaN), 'fit');
  assert.equal(C.normalizeZoom('1.25'), 1.25);
  assert.equal(C.normalizeZoom(1.23456), 1.235);
});

test('stepZoom walks the preset steps both ways and stops at the ends', () => {
  assert.equal(C.stepZoom(1, 1), 1.1);
  assert.equal(C.stepZoom(1, -1), 0.9);
  assert.equal(C.stepZoom(0.63, 1), 0.67);
  assert.equal(C.stepZoom(0.63, -1), 0.5);
  assert.equal(C.stepZoom(4, 1), 4);
  assert.equal(C.stepZoom(0.1, -1), 0.1);
  let s = 0.1;
  for (let i = 0; i < 40; i++) s = C.stepZoom(s, 1);
  assert.equal(s, C.ZOOM_MAX);
});

test('wheelZoom: wheel up zooms in, down zooms out, big deltas are capped', () => {
  assert.ok(C.wheelZoom(1, -100) > 1);
  assert.ok(C.wheelZoom(1, 100) < 1);
  assert.equal(C.wheelZoom(1, -100000), C.wheelZoom(1, -100));
  assert.equal(C.wheelZoom(4, -100), 4);
  assert.equal(C.wheelZoom(0.1, 100), 0.1);
  assert.equal(C.wheelZoom(1, 0), 1);
});

test('zoomLabel names the mode and shows the real percent', () => {
  assert.equal(C.zoomLabel('fit', 0.634), 'Fit (63%)');
  assert.equal(C.zoomLabel('width', 1.2), 'Width (120%)');
  assert.equal(C.zoomLabel(1.5, 1.5), '150%');
});

test('slideAtScroll picks the slide at 40% down the view, and normalizeView', () => {
  const tops = [0, 600, 1200, 1800];
  assert.equal(C.slideAtScroll(tops, 0, 500), 1);
  assert.equal(C.slideAtScroll(tops, 400, 500), 2); // line at 600
  assert.equal(C.slideAtScroll(tops, 399, 500), 1);
  assert.equal(C.slideAtScroll(tops, 5000, 500), 4);
  assert.equal(C.slideAtScroll([], 0, 500), 1);
  assert.equal(C.normalizeView('scroll'), 'scroll');
  assert.equal(C.normalizeView('grid'), 'slide');
  assert.equal(C.normalizeView(undefined), 'slide');
});

// ---------- protocol ----------
test('viewPrefs message: valid views and zooms pass, everything else is dropped', () => {
  assert.deepEqual(parseWebviewMessage({ type: 'viewPrefs', view: 'scroll', zoom: 'width', x: 1 }), { type: 'viewPrefs', view: 'scroll', zoom: 'width' });
  assert.deepEqual(parseWebviewMessage({ type: 'viewPrefs', view: 'slide', zoom: 1.5 }), { type: 'viewPrefs', view: 'slide', zoom: 1.5 });
  for (const m of [{ type: 'viewPrefs', view: 'grid', zoom: 'fit' }, { type: 'viewPrefs', view: 'slide', zoom: 5 },
    { type: 'viewPrefs', view: 'slide', zoom: 0 }, { type: 'viewPrefs', view: 'slide', zoom: '1' }, { type: 'viewPrefs', view: 'slide', zoom: Infinity },
    { type: 'viewPrefs', view: 'slide' }, { type: 'viewPrefs', zoom: 'fit' }]) {
    assert.equal(parseWebviewMessage(m), null, JSON.stringify(m));
  }
});

test('config: defaultView and defaultZoom with safe fallbacks', () => {
  const d = readConfig((k, def) => def);
  assert.equal(d.defaultView, 'slide');
  assert.equal(d.defaultZoom, 'fit');
  const s = readConfig((k, def) => ({ defaultView: 'scroll', defaultZoom: 'width' }[k] ?? def));
  assert.equal(s.defaultView, 'scroll');
  assert.equal(s.defaultZoom, 'width');
  const bad = readConfig((k, def) => ({ defaultView: 'x', defaultZoom: 7 }[k] ?? def));
  assert.equal(bad.defaultView, 'slide');
  assert.equal(bad.defaultZoom, 'fit');
});

// ---------- package.json and webview ----------
test('package.json: zoom and scroll commands, palette rules, keybindings and settings', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  const c = pkg.contributes;
  const ids = ['zoomIn', 'zoomOut', 'zoomFit', 'zoomWidth', 'zoomActual', 'toggleScrollView'].map((n) => 'pptxViewer.' + n);
  for (const id of ids) {
    assert.ok(c.commands.some((x) => x.command === id), id);
    const rule = c.menus.commandPalette.find((x) => x.command === id);
    assert.equal(rule && rule.when, 'activeCustomEditorId == pptxViewer.preview', id);
  }
  const kb = (cmd) => c.keybindings.filter((k) => k.command === cmd);
  assert.ok(kb('pptxViewer.zoomIn').some((k) => k.mac === 'cmd+=' && k.key === 'ctrl+='));
  assert.ok(kb('pptxViewer.zoomOut').some((k) => k.mac === 'cmd+-' && k.key === 'ctrl+-'));
  for (const k of [...kb('pptxViewer.zoomIn'), ...kb('pptxViewer.zoomOut')]) assert.match(k.when, /^activeCustomEditorId == pptxViewer\.preview && /);
  const props = c.configuration.properties;
  assert.deepEqual(props['pptxViewer.defaultView'].enum, ['slide', 'scroll']);
  assert.equal(props['pptxViewer.defaultView'].default, 'slide');
  assert.deepEqual(props['pptxViewer.defaultZoom'].enum, ['fit', 'width']);
  assert.equal(props['pptxViewer.defaultZoom'].default, 'fit');
});

test('webview: toolbar has the zoom controls and Scroll button, and viewer.js handles every forwarded command', () => {
  const html = buildHtml({ css: 'c', core: 'k', viewer: 'v', preview: 'p' }, 'vscode-webview://x', 'n');
  for (const id of ['zoomOutBtn', 'zoomSelect', 'zoomInBtn', 'scrollBtn']) assert.match(html, new RegExp(`id="${id}"`), id);
  for (const v of ['fit', 'width', '0.5', '1', '2', '4']) assert.match(html, new RegExp(`<option value="${v.replace('.', '\\.')}"`), v);
  const viewer = fs.readFileSync(path.join(ROOT, 'media/viewer.js'), 'utf8');
  const { WEBVIEW_COMMANDS } = loadExtension();
  for (const name of WEBVIEW_COMMANDS) assert.ok(viewer.includes(`m.name === '${name}'`), name);
  // Deck text in the scroll view goes in with textContent only.
  assert.ok(!/innerHTML|insertAdjacentHTML|outerHTML\s*=/.test(viewer));
});

// ---------- extension host with a stub of the vscode API ----------
let stubbed = null;
function loadExtension() {
  if (stubbed) return stubbed;
  const commands = new Map();
  const uri = (p) => ({ path: p, fsPath: p, toString: () => 'file://' + p });
  const vscodeStub = {
    Uri: { joinPath: (u, ...p) => uri(path.posix.join(u.path, ...p)), parse: (s) => uri(s) },
    RelativePattern: class { constructor(b, p) { this.base = b; this.pattern = p; } },
    workspace: {
      getConfiguration: () => ({ get: (k, d) => d }),
      fs: { stat: async (u) => ({ size: fs.statSync(u.path).size }), readFile: async (u) => fs.readFileSync(u.path) },
      createFileSystemWatcher: () => ({ onDidChange: () => ({ dispose() {} }), onDidCreate: () => ({ dispose() {} }), dispose() {} }),
      onDidChangeConfiguration: () => ({ dispose() {} }),
    },
    window: { registerCustomEditorProvider: () => ({ dispose() {} }), setStatusBarMessage: () => ({ dispose() {} }), showOpenDialog: async () => undefined },
    env: { clipboard: { writeText: async () => {} } },
    commands: { registerCommand: (id, fn) => { commands.set(id, fn); return { dispose() {} }; }, executeCommand: async () => {} },
  };
  const origLoad = Module._load;
  Module._load = function (req, ...rest) { return req === 'vscode' ? vscodeStub : origLoad.call(this, req, ...rest); };
  const ext = require('../src/extension');
  stubbed = Object.assign({ commands, uri }, ext);
  return stubbed;
}

function memento() {
  const data = new Map();
  return { get: (k, d) => (data.has(k) ? data.get(k) : d), update: async (k, v) => { data.set(k, v); }, data };
}

function openPanel(provider, uri, file) {
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
  provider.resolveCustomEditor(provider.openCustomDocument(uri(path.join(FIX, file))), panel);
  return { panel, sent, post: (m) => onMsg(m) };
}
const wait = () => new Promise((r) => setTimeout(r, 30));

test('host: the zoom and scroll commands are registered and forwarded to the active viewer', async () => {
  const ext = loadExtension();
  const ctx = { subscriptions: [], extensionUri: ext.uri('/ext'), workspaceState: memento(), globalState: memento() };
  const { provider } = ext.activate(ctx);
  for (const n of ext.WEBVIEW_COMMANDS) assert.ok(ext.commands.has('pptxViewer.' + n), n);
  const p = openPanel(provider, ext.uri, 'quarterly-review.pptx');
  await ext.commands.get('pptxViewer.zoomIn')();
  await ext.commands.get('pptxViewer.toggleScrollView')();
  assert.deepEqual(p.sent.filter((m) => m.type === 'command').map((m) => m.name), ['zoomIn', 'toggleScrollView']);
});

test('host: the view and zoom are kept per file and sent back in init on the next open', async () => {
  const ext = loadExtension();
  const ctx = { extensionUri: ext.uri('/ext'), workspaceState: memento() };
  const provider = new ext.PptxEditorProvider(ctx);
  const a = openPanel(provider, ext.uri, 'quarterly-review.pptx');
  a.post({ type: 'ready' });
  await wait();
  assert.equal(a.sent[0].type, 'init');
  assert.equal(a.sent[0].prefs, null, 'no saved choice the first time');
  await a.post({ type: 'viewPrefs', view: 'scroll', zoom: 1.5 });
  await a.post({ type: 'viewPrefs', view: 'nope', zoom: 1 }); // dropped by the protocol
  const b = openPanel(provider, ext.uri, 'quarterly-review.pptx');
  b.post({ type: 'ready' });
  await wait();
  assert.deepEqual(b.sent[0].prefs, { view: 'scroll', zoom: 1.5 });
  const other = openPanel(provider, ext.uri, 'many-slides.pptx');
  other.post({ type: 'ready' });
  await wait();
  assert.equal(other.sent[0].prefs, null, 'another file keeps its own choice');
});

test('host: only the newest 200 files are kept', async () => {
  const ext = loadExtension();
  const ws = memento();
  const provider = new ext.PptxEditorProvider({ extensionUri: ext.uri('/ext'), workspaceState: ws });
  for (let i = 0; i < 205; i++) await provider.savePrefs(ext.uri('/deck' + i + '.pptx'), 'slide', 'width');
  const all = ws.get('pptxViewer.viewPrefs');
  assert.equal(Object.keys(all).length, 200);
  assert.ok(provider.getPrefs(ext.uri('/deck204.pptx')));
  assert.equal(provider.getPrefs(ext.uri('/deck0.pptx')), null);
});
