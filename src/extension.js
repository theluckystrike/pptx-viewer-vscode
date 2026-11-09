'use strict';
const vscode = require('vscode');
const path = require('path');
const fs = require('fs');
const { buildHtml } = require('./html');
const { parseWebviewMessage, readConfig } = require('./protocol');
const { parseDeck, slideToMarkdown, deckToMarkdown } = require('./pptx');
const { parsePpt, sniffPpt } = require('./ppt');
const { createRatingPrompt } = require('./rate');
const { registerSupport, commandUris, showLinks } = require('./support');
const { createPro } = require('./pro');
const { createExportTools } = require('./exportTools');

let rating = { success() {}, failure() {} };

const VIEW_TYPE = 'pptxViewer.preview';
const CONFIG = 'pptxViewer';
const PREFS_KEY = 'pptxViewer.viewPrefs';
const PREFS_MAX = 200;
const WEBVIEW_COMMANDS = ['find', 'toggleNotes', 'zoomIn', 'zoomOut', 'zoomFit', 'zoomWidth', 'zoomActual', 'toggleScrollView', 'present', 'copySlidePng', 'saveSlidePng'];
const LEGACY_EXT = /\.(ppt|pps|pot)$/i;
// The slide renderer in the webview unpacks the XML again and needs about 40x its size in memory.
// A deck over these limits is shown as text only (real decks: largest part well under 1 MB).
const DRAW_MAX_PART_BYTES = 16 * 1024 * 1024;
const DRAW_MAX_TOTAL_BYTES = 96 * 1024 * 1024;
const EMPTY_FILE_MESSAGE = 'This file is empty (0 bytes). It has no slides. If it is still being written or downloaded, it opens by itself when it is done.';

/**
 * Deck text and whether the slides can be drawn. Old .ppt files (and a .pptx saved with a .ppt name)
 * are told apart by their first bytes, not by the name.
 */
function readDeck(data, fileName) {
  if (!data || data.length === 0) throw new Error(EMPTY_FILE_MESSAGE);
  const kind = sniffPpt(data);
  if (kind === 'ole') return { deck: parsePpt(data), drawable: false };
  if (kind === 'other' && LEGACY_EXT.test(fileName)) throw new Error('This file has a PowerPoint name, but it is neither a PowerPoint 97-2003 file nor a .pptx deck.');
  const deck = parseDeck(data);
  const st = deck.xmlStats || { maxPart: 0, total: 0 };
  if (st.maxPart > DRAW_MAX_PART_BYTES || st.total > DRAW_MAX_TOTAL_BYTES) {
    const mb = Math.round(Math.max(st.maxPart, st.total) / (1024 * 1024));
    deck.textOnly = `This deck has ${mb} MB of slide XML when unpacked, which is too much to draw safely. Text, notes and search still work. Open it in PowerPoint, Keynote or LibreOffice to see the slides.`;
    return { deck, drawable: false };
  }
  return { deck, drawable: true };
}

/** Calls onChange when a local file's size or mtime changes (1 s stat poll). Returns a disposable. */
function pollLocalFile(uri, onChange, interval = 1000) {
  if (!uri || uri.scheme !== 'file' || !uri.fsPath) return { dispose() {} };
  const file = uri.fsPath;
  const listener = (cur, prev) => {
    if (cur.mtimeMs !== prev.mtimeMs || cur.size !== prev.size) onChange();
  };
  try { fs.watchFile(file, { interval, persistent: false }, listener); } catch { return { dispose() {} }; }
  return { dispose() { fs.unwatchFile(file, listener); } };
}

// Node Buffers are serialised to JSON by postMessage; a plain Uint8Array is sent as bytes.
function toPlainBytes(data) {
  if (Object.getPrototypeOf(data) === Uint8Array.prototype && data.byteOffset === 0 && data.byteLength === data.buffer.byteLength) return data;
  const out = new Uint8Array(data.byteLength);
  out.set(data);
  return out;
}

class DeckDocument {
  constructor(uri) { this.uri = uri; }
  dispose() {}
}

class PptxEditorProvider {
  constructor(context) {
    this.context = context;
    this.active = null;
    this.entries = new Set();
  }

  /** What the webview shows of Zovo Pro: the Export button (show) and its PRO badge (badge). */
  proState() {
    const p = this.pro;
    if (!p) return { show: false, badge: false };
    return { show: p.shown(), badge: !p.isPro() };
  }

  postPro() {
    const st = this.proState();
    for (const e of this.entries) if (e.ready) e.panel.webview.postMessage({ type: 'pro', pro: st });
  }

  openCustomDocument(uri) { return new DeckDocument(uri); }

  // View and zoom per file, kept in workspaceState for the newest PREFS_MAX files.
  getPrefs(uri) {
    const store = this.context.workspaceState;
    const all = (store && store.get(PREFS_KEY)) || {};
    const p = all[uri.toString()];
    return p && typeof p === 'object' ? { view: p.view, zoom: p.zoom } : null;
  }

  async savePrefs(uri, view, zoom) {
    const store = this.context.workspaceState;
    if (!store) return;
    const all = Object.assign({}, store.get(PREFS_KEY) || {});
    all[uri.toString()] = { view, zoom, t: Date.now() };
    const keys = Object.keys(all);
    if (keys.length > PREFS_MAX) {
      keys.sort((a, b) => (all[a].t || 0) - (all[b].t || 0));
      for (const k of keys.slice(0, keys.length - PREFS_MAX)) delete all[k];
    }
    await store.update(PREFS_KEY, all);
  }

  config() {
    const c = vscode.workspace.getConfiguration(CONFIG);
    return readConfig((k, d) => c.get(k, d));
  }

  async resolveCustomEditor(document, panel) {
    const media = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    const webview = panel.webview;
    // Command links are limited to the two support commands and the settings page (footer "Hide" link).
    webview.options = { enableScripts: true, localResourceRoots: [media], enableCommandUris: commandUris('pptxViewer') };
    const u = (...p) => webview.asWebviewUri(vscode.Uri.joinPath(media, ...p)).toString();
    webview.html = buildHtml({ css: u('viewer.css'), core: u('core.js'), viewer: u('viewer.js'), preview: u('vendor', 'pptx-preview.umd.js') }, webview.cspSource, undefined, { showSupportLinks: showLinks(vscode, 'pptxViewer') });

    const fileName = path.posix.basename(document.uri.path);
    const entry = { panel, uri: document.uri, deck: null, slide: 1, forceLoad: false, ready: false };
    if (panel.active) this.active = entry;
    this.entries.add(entry);
    const disposables = [];
    const send = (m) => webview.postMessage(m);
    let disposed = false;

    const load = async () => {
      try {
        const cfg = this.config();
        const stat = await vscode.workspace.fs.stat(document.uri);
        if (stat.size > cfg.maxFileSizeBytes && !entry.forceLoad) {
          send({ type: 'tooLarge', size: stat.size, limit: cfg.maxFileSizeBytes });
          return;
        }
        const data = await vscode.workspace.fs.readFile(document.uri);
        if (disposed) return;
        // Text first: the outline, notes and search work before any slide is drawn.
        const { deck, drawable } = readDeck(data, fileName);
        entry.deck = deck;
        send({ type: 'deck', deck: entry.deck, fileName });
        // Rating: a deck counts once its slides are drawn (webview 'rendered'), or now if no slide is drawn.
        entry.waitRender = !!(cfg.renderSlides && drawable);
        if (!entry.waitRender && !entry.counted) { entry.counted = true; rating.success(); }
        if (cfg.renderSlides && drawable) send({ type: 'bytes', data: toPlainBytes(data) });
      } catch (e) {
        rating.failure();
        send({ type: 'error', message: e && e.message ? e.message : String(e) });
      }
    };

    disposables.push(webview.onDidReceiveMessage(async (raw) => {
      const m = parseWebviewMessage(raw);
      if (!m) return;
      if (m.type === 'ready') {
        entry.ready = true;
        send({ type: 'init', config: this.config(), fileName, prefs: this.getPrefs(document.uri), pro: this.proState() });
        load();
      } else if (m.type === 'openInDefaultApp') {
        await vscode.env.openExternal(document.uri);
      } else if (m.type === 'present') {
        await this.present(entry, m.on);
      } else if (m.type === 'loadAnyway') {
        entry.forceLoad = true;
        load();
      } else if (m.type === 'slideChanged') {
        entry.slide = m.slide;
      } else if (m.type === 'copySlideMarkdown') {
        await this.copySlide(entry, m.slide);
      } else if (m.type === 'copyDeckMarkdown') {
        await this.copyDeck(entry);
      } else if (m.type === 'viewPrefs') {
        await this.savePrefs(document.uri, m.view, m.zoom);
      } else if (m.type === 'savePng') {
        await this.savePng(entry, m.slide, m.data);
      } else if (m.type === 'pngDone') {
        vscode.window.setStatusBarMessage(`$(file-media) Slide ${m.slide} copied as a picture`, 3000);
      } else if (m.type === 'pngFailed') {
        vscode.window.showWarningMessage('PPTX Viewer: ' + m.message);
      } else if (m.type === 'exportMenu') {
        if (this.tools) { this.active = entry; this.tools.menu(); }
      } else if (m.type === 'exportSlide' || m.type === 'exportDone') {
        if (this.tools) this.tools.onMessage(entry, m);
      } else if (m.type === 'rendered') {
        entry.rendered = m;
        if (entry.waitRender && !entry.counted) { entry.counted = true; rating.success(); }
      } else if (m.type === 'error') {
        entry.lastError = m.message;
        rating.failure();
      }
    }));

    let timer = null;
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(document.uri, '..'), fileName));
    const onChange = () => { clearTimeout(timer); timer = setTimeout(() => { if (entry.ready && !disposed) load(); }, 400); };
    disposables.push(watcher, watcher.onDidChange(onChange), watcher.onDidCreate(onChange));
    // VS Code's watcher can miss writes made by other programs (a script or an AI agent that saves the
    // deck again). For local files, a 1 s stat poll catches them as well.
    disposables.push(pollLocalFile(document.uri, onChange));
    disposables.push(vscode.workspace.onDidChangeConfiguration((e) => {
      if (e.affectsConfiguration(CONFIG) && entry.ready) send({ type: 'config', config: this.config() });
    }));
    disposables.push(panel.onDidChangeViewState(() => {
      if (panel.active) this.active = entry;
      else if (this.active === entry) this.active = null;
    }));
    panel.onDidDispose(() => {
      disposed = true;
      this.entries.delete(entry);
      if (this.tools) this.tools.onPanelClosed(entry);
      if (entry.zen) { entry.zen = false; vscode.commands.executeCommand('workbench.action.toggleZenMode').then(undefined, () => {}); }
      clearTimeout(timer);
      if (this.active === entry) this.active = null;
      disposables.forEach((d) => d.dispose());
    });
  }

  async copySlide(entry, n) {
    if (!entry || !entry.deck) return;
    const slide = entry.deck.slides[(n || entry.slide) - 1];
    if (!slide) return;
    await vscode.env.clipboard.writeText(slideToMarkdown(slide));
    vscode.window.setStatusBarMessage(`$(markdown) Slide ${slide.index} copied as Markdown`, 3000);
  }

  async copyDeck(entry) {
    if (!entry || !entry.deck) return;
    await vscode.env.clipboard.writeText(deckToMarkdown(entry.deck, path.posix.basename(entry.uri.path)));
    vscode.window.setStatusBarMessage(`$(markdown) ${entry.deck.slides.length} slides copied as Markdown`, 3000);
  }

  /** Asks where to save the slide picture (next to the deck by default) and writes it. */
  async savePng(entry, n, data) {
    const base = path.posix.basename(entry.uri.path).replace(/\.[^.]+$/, '');
    const defaultUri = vscode.Uri.joinPath(entry.uri, '..', `${base}-slide-${n}.png`);
    const target = await vscode.window.showSaveDialog({ defaultUri, filters: { 'PNG image': ['png'] }, saveLabel: 'Save Slide' });
    if (!target) return;
    await vscode.workspace.fs.writeFile(target, data);
    vscode.window.setStatusBarMessage(`$(file-media) Slide ${n} saved as ${path.posix.basename(target.path)}`, 4000);
  }

  /** Present mode: the webview shows one slide on black; VS Code goes to Zen Mode (full screen) while it lasts. */
  async present(entry, on) {
    if (!this.config().presentFullScreen || !!entry.zen === on) return;
    entry.zen = on;
    try { await vscode.commands.executeCommand('workbench.action.toggleZenMode'); } catch { entry.zen = false; }
  }

  sendCommand(name) {
    if (this.active) this.active.panel.webview.postMessage({ type: 'command', name });
  }
}

function activate(context) {
  rating = createRatingPrompt({ vscode, context, extensionId: 'zovo.pptx-slide-viewer', section: 'pptxViewer', name: 'PPTX Viewer' });
  registerSupport({ vscode, context, prefix: 'pptxViewer', selfId: 'zovo.pptx-slide-viewer' });
  const provider = new PptxEditorProvider(context);
  // Zovo Pro (src/pro.js): all slides as PNG, the deck as PDF, decks in a folder to Markdown. Local checks only.
  const pro = createPro({ vscode, context });
  provider.pro = pro;
  provider.tools = createExportTools({ vscode, provider, pro, readDeck });
  pro.ready.then(() => provider.postPro(), () => {});
  context.subscriptions.push(pro.onDidChange(() => provider.postPro()));
  context.subscriptions.push(vscode.workspace.onDidChangeConfiguration((e) => { if (e.affectsConfiguration('pptxViewer.showProFeatures')) provider.postPro(); }));
  if (vscode.window.registerUriHandler) context.subscriptions.push(vscode.window.registerUriHandler({ handleUri: (uri) => pro.license.handleUri(uri) }));
  context.subscriptions.push(
    vscode.commands.registerCommand('pptxViewer.exportMenu', () => provider.tools.menu()),
    vscode.commands.registerCommand('pptxViewer.exportSlidesPng', () => provider.tools.exportPngs()),
    vscode.commands.registerCommand('pptxViewer.exportPdf', () => provider.tools.exportPdf()),
    vscode.commands.registerCommand('pptxViewer.decksToMarkdown', (uri) => provider.tools.markdownBatch(uri instanceof vscode.Uri ? uri : undefined)),
  );
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.commands.registerCommand('pptxViewer.copySlideMarkdown', () => provider.copySlide(provider.active)),
    vscode.commands.registerCommand('pptxViewer.copyDeckMarkdown', () => provider.copyDeck(provider.active)),
    ...WEBVIEW_COMMANDS.map((name) => vscode.commands.registerCommand('pptxViewer.' + name, () => provider.sendCommand(name))),
    vscode.commands.registerCommand('pptxViewer.openWith', async (uri) => {
      const target = uri instanceof vscode.Uri ? uri
        : ((await vscode.window.showOpenDialog({ filters: { PowerPoint: ['pptx', 'ppsx', 'potx', 'pptm', 'ppt', 'pps', 'pot'] }, canSelectMany: false })) || [])[0];
      if (target) await vscode.commands.executeCommand('vscode.openWith', target, VIEW_TYPE);
    }),
  );
  return { provider, pro };
}

function deactivate() {}

module.exports = { activate, deactivate, PptxEditorProvider, VIEW_TYPE, toPlainBytes, WEBVIEW_COMMANDS, readDeck, EMPTY_FILE_MESSAGE, pollLocalFile };
