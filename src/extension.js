'use strict';
const vscode = require('vscode');
const path = require('path');
const { buildHtml } = require('./html');
const { parseWebviewMessage, readConfig } = require('./protocol');
const { parseDeck, slideToMarkdown, deckToMarkdown } = require('./pptx');
const { createRatingPrompt } = require('./rate');
const { registerSupport, commandUris, showLinks } = require('./support');

let rating = { success() {} };

const VIEW_TYPE = 'pptxViewer.preview';
const CONFIG = 'pptxViewer';

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
  }

  openCustomDocument(uri) { return new DeckDocument(uri); }

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
        entry.deck = parseDeck(data);
        send({ type: 'deck', deck: entry.deck, fileName });
        if (!entry.counted) { entry.counted = true; rating.success(); }
        if (cfg.renderSlides) send({ type: 'bytes', data: toPlainBytes(data) });
      } catch (e) {
        send({ type: 'error', message: e && e.message ? e.message : String(e) });
      }
    };

    disposables.push(webview.onDidReceiveMessage(async (raw) => {
      const m = parseWebviewMessage(raw);
      if (!m) return;
      if (m.type === 'ready') {
        entry.ready = true;
        send({ type: 'init', config: this.config(), fileName });
        load();
      } else if (m.type === 'loadAnyway') {
        entry.forceLoad = true;
        load();
      } else if (m.type === 'slideChanged') {
        entry.slide = m.slide;
      } else if (m.type === 'copySlideMarkdown') {
        await this.copySlide(entry, m.slide);
      } else if (m.type === 'copyDeckMarkdown') {
        await this.copyDeck(entry);
      } else if (m.type === 'rendered') {
        entry.rendered = m;
      } else if (m.type === 'error') {
        entry.lastError = m.message;
      }
    }));

    let timer = null;
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(document.uri, '..'), fileName));
    const onChange = () => { clearTimeout(timer); timer = setTimeout(() => { if (entry.ready) load(); }, 400); };
    disposables.push(watcher, watcher.onDidChange(onChange), watcher.onDidCreate(onChange));
    disposables.push(vscode.workspace.onDidChangeConfiguration((e) => {
      if (e.affectsConfiguration(CONFIG) && entry.ready) send({ type: 'config', config: this.config() });
    }));
    disposables.push(panel.onDidChangeViewState(() => {
      if (panel.active) this.active = entry;
      else if (this.active === entry) this.active = null;
    }));
    panel.onDidDispose(() => {
      disposed = true;
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

  sendCommand(name) {
    if (this.active) this.active.panel.webview.postMessage({ type: 'command', name });
  }
}

function activate(context) {
  rating = createRatingPrompt({ vscode, context, extensionId: 'zovo.pptx-slide-viewer', section: 'pptxViewer', message: (n) => `PPTX Viewer has opened ${n} files for you. A rating helps other people find it.` });
  registerSupport({ vscode, context, prefix: 'pptxViewer', selfId: 'zovo.pptx-slide-viewer' });
  const provider = new PptxEditorProvider(context);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.commands.registerCommand('pptxViewer.copySlideMarkdown', () => provider.copySlide(provider.active)),
    vscode.commands.registerCommand('pptxViewer.copyDeckMarkdown', () => provider.copyDeck(provider.active)),
    vscode.commands.registerCommand('pptxViewer.find', () => provider.sendCommand('find')),
    vscode.commands.registerCommand('pptxViewer.toggleNotes', () => provider.sendCommand('toggleNotes')),
    vscode.commands.registerCommand('pptxViewer.openWith', async (uri) => {
      const target = uri instanceof vscode.Uri ? uri
        : ((await vscode.window.showOpenDialog({ filters: { PowerPoint: ['pptx', 'ppsx', 'potx', 'pptm'] }, canSelectMany: false })) || [])[0];
      if (target) await vscode.commands.executeCommand('vscode.openWith', target, VIEW_TYPE);
    }),
  );
  return { provider };
}

function deactivate() {}

module.exports = { activate, deactivate, PptxEditorProvider, VIEW_TYPE, toPlainBytes };
