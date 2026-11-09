'use strict';
// Zovo Pro in PPTX Viewer: the PDF writer, file names, the combined Markdown, the export protocol, the free-try
// rules, the host flows with a fake vscode (slides answered by a stand-in webview) and activation. No network.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const os = require('os');
const Module = require('module');
const { jpegInfo, imagesToPdf, pdfString, slideSizePt } = require('../src/pdfwriter');
const { slideFileName, deckBase, combinedMarkdown } = require('../src/exportTools');
const { parseWebviewMessage } = require('../src/protocol');

const FIX = path.join(__dirname, 'fixtures');
const JPG = fs.readFileSync(path.join(FIX, 'slide.jpg'));
const PNG = fs.readFileSync(path.join(FIX, 'slide.png'));

// ---------- PDF writer ----------
test('jpegInfo: reads the frame size of a real JPEG, refuses other bytes', () => {
  assert.deepEqual(jpegInfo(JPG), { width: 64, height: 64, comps: 3 });
  assert.equal(jpegInfo(PNG), null);
  assert.equal(jpegInfo(new Uint8Array([0xff, 0xd8, 0xff, 0xda, 0, 2])), null);
  assert.equal(jpegInfo(JPG.subarray(0, 30)), null);
  assert.equal(jpegInfo(null), null);
});

test('imagesToPdf: one page per slide at the slide size, a correct xref table, title in Info', () => {
  const size = slideSizePt({ widthPx: 1280, heightPx: 720 });
  assert.deepEqual(size, { widthPt: 960, heightPt: 540 });
  const pdf = imagesToPdf([{ jpeg: JPG }, { jpeg: JPG }, { jpeg: JPG }], Object.assign({ title: 'Q3 review (draft)' }, size));
  const s = pdf.toString('latin1');
  assert.ok(s.startsWith('%PDF-1.4\n'));
  assert.ok(s.endsWith('%%EOF\n'));
  assert.match(s, /\/Type \/Pages \/Count 3 \/Kids \[4 0 R 7 0 R 10 0 R\]/);
  assert.equal((s.match(/\/MediaBox \[0 0 960 540\]/g) || []).length, 3);
  assert.match(s, /\/Title \(Q3 review \\\(draft\\\)\)/);
  assert.equal((s.match(/\/Filter \/DCTDecode \/Length 3190/g) || []).length, 3);
  // Every xref offset points at its "n 0 obj".
  const startxref = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(s)[1]);
  const table = s.slice(startxref).split('\n');
  assert.equal(table[0], 'xref');
  const count = Number(table[1].split(' ')[1]);
  assert.equal(count, 13);
  for (let n = 1; n < count; n++) {
    const off = Number(table[2 + n].slice(0, 10));
    assert.equal(s.slice(off, off + String(n).length + 6), n + ' 0 obj', 'object ' + n);
  }
  assert.throws(() => imagesToPdf([{ jpeg: PNG }], size), /Slide 1 is not a JPEG/);
  assert.throws(() => imagesToPdf([], size), /No slides/);
});

test('imagesToPdf: pdf.js opens it with the right page count and size (when pdf.js is on this machine)', async (t) => {
  const pdfjsPath = process.env.PDFJS_LEGACY || path.join(os.homedir(), 'plugin-market-intel-work/factory/products/pdf-viewer-vscode/node_modules/pdfjs-dist/legacy/build/pdf.mjs');
  if (!fs.existsSync(pdfjsPath)) return t.skip('pdfjs-dist not found');
  globalThis.DOMMatrix ??= class { constructor() { Object.assign(this, { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }); } };
  globalThis.Path2D ??= class {};
  globalThis.ImageData ??= class {};
  const pdfjs = await import(pdfjsPath);
  pdfjs.GlobalWorkerOptions.workerSrc = path.join(path.dirname(pdfjsPath), 'pdf.worker.mjs');
  const pdf = imagesToPdf([{ jpeg: JPG }, { jpeg: JPG }], { widthPt: 720, heightPt: 405, title: 'Café ✓' });
  const doc = await pdfjs.getDocument({ data: new Uint8Array(pdf), isEvalSupported: false, verbosity: 0 }).promise;
  assert.equal(doc.numPages, 2);
  const page = await doc.getPage(2);
  assert.deepEqual(page.view, [0, 0, 720, 405]);
  const ops = await page.getOperatorList();
  assert.ok(ops.fnArray.includes(pdfjs.OPS.paintImageXObject), 'the slide picture is painted');
  const meta = await doc.getMetadata();
  assert.equal(meta.info.Title, 'Café ✓');
  await doc.destroy();
});

test('pdfString: ASCII literal with escapes, other text as UTF-16BE hex', () => {
  assert.equal(pdfString('a(b)\\c'), '(a\\(b\\)\\\\c)');
  assert.equal(pdfString('é'), '<FEFF00E9>');
  assert.equal(pdfString('😀'), '<FEFFD83DDE00>');
});

// ---------- names and Markdown ----------
test('file names sort in slide order; deck base name', () => {
  assert.equal(slideFileName('deck', 7, 12, 'png'), 'deck-slide-07.png');
  assert.equal(slideFileName('deck', 7, 120, 'png'), 'deck-slide-007.png');
  assert.equal(slideFileName('deck', 1, 3, 'png'), 'deck-slide-01.png');
  assert.equal(deckBase('Q3 review.final.pptx'), 'Q3 review.final');
  assert.equal(deckBase('.pptx'), 'deck');
});

test('combinedMarkdown: one heading per deck path, slides one level down, failures listed', () => {
  const md = combinedMarkdown([
    { rel: 'a/q3.pptx', markdown: '# q3.pptx\n\n## Revenue\n\n- up\n\n## Costs\n\n- down\n', slides: 2 },
    { rel: 'b/broken.ppt', error: 'not a deck' },
  ]);
  assert.match(md, /^# Slides\n\n1 deck, 2 slides\.\n\n## a\/q3\.pptx\n\n### Revenue\n\n- up\n\n### Costs\n\n- down\n/);
  assert.match(md, /## Not converted\n\n- b\/broken\.ppt: not a deck/);
});

// ---------- protocol ----------
test('protocol: export messages are checked', () => {
  assert.deepEqual(parseWebviewMessage({ type: 'exportMenu', x: 1 }), { type: 'exportMenu' });
  const s = parseWebviewMessage({ type: 'exportSlide', id: 2, slide: 3, data: new Uint8Array(JPG) });
  assert.equal(s.slide, 3);
  assert.ok(s.data instanceof Uint8Array);
  assert.equal(parseWebviewMessage({ type: 'exportSlide', id: 2, slide: 3, data: 'x' }), null);
  assert.equal(parseWebviewMessage({ type: 'exportSlide', id: 0, slide: 3, data: new Uint8Array(JPG) }), null);
  assert.deepEqual(parseWebviewMessage({ type: 'exportDone', id: 2, failed: [4, 'x', -1], cancelled: true }), { type: 'exportDone', id: 2, failed: [4], cancelled: true });
  assert.equal(parseWebviewMessage({ type: 'exportDone', id: 2, error: 'e'.repeat(500) }).error.length, 300);
});

// ---------- host flows with a fake vscode ----------
function fakeVscode(o) {
  o = o || {};
  const log = { info: [], warn: [], modal: [], status: [], picks: [], saves: [], opens: [], written: new Map(), commands: [], external: [], shown: [] };
  const answers = { pick: o.pick || [], save: o.save || [], open: o.open || [], info: o.info || [], warn: o.warn || [] };
  class Uri {
    constructor(scheme, p) { this.scheme = scheme; this.path = p; this.fsPath = p; }
    static file(p) { return new Uri('file', p); }
    static joinPath(u, ...parts) { return new Uri(u.scheme, path.posix.join(u.path, ...parts)); }
    with(c) { return new Uri(c.scheme ?? this.scheme, c.path ?? this.path); }
    toString() { return this.scheme + '://' + this.path; }
  }
  const settings = Object.assign({ showProFeatures: true }, o.settings || {});
  const state = new Map();
  const realStat = (p) => { try { return fs.statSync(p); } catch { return null; } };
  const vscode = {
    Uri,
    ProgressLocation: { Notification: 15 },
    window: {
      showInformationMessage: async (msg, ...rest) => {
        if (rest[0] && typeof rest[0] === 'object' && rest[0].modal) { log.modal.push(msg); return answers.info.shift(); }
        log.info.push(msg); return answers.info.shift();
      },
      showWarningMessage: async (msg, ...rest) => { log.warn.push(msg); return rest[0] && rest[0].modal ? answers.warn.shift() : undefined; },
      showQuickPick: async (items, b) => { log.picks.push({ items, b }); const want = answers.pick.shift(); return want == null ? undefined : items.find((i) => i.id === want); },
      showSaveDialog: async (b) => { log.saves.push(b); const v = answers.save.shift(); return v === 'default' ? b.defaultUri : v; },
      showOpenDialog: async (b) => { log.opens.push(b); return answers.open.shift(); },
      setStatusBarMessage: (m) => { log.status.push(m); return { dispose() {} }; },
      withProgress: async (_o, fn) => fn({ report() {} }, { isCancellationRequested: false, onCancellationRequested: () => ({ dispose() {} }) }),
      showTextDocument: async (d) => { log.shown.push(d); },
    },
    workspace: {
      workspaceFolders: undefined,
      getConfiguration: () => ({ get: (k, d) => (k in settings ? settings[k] : d) }),
      fs: {
        stat: async (u) => { if (log.written.has(u.path)) return { size: log.written.get(u.path).length }; const st = realStat(u.path); if (!st) throw new Error('ENOENT'); return { size: st.size }; },
        readFile: async (u) => fs.readFileSync(u.path),
        writeFile: async (u, b) => { log.written.set(u.path, Buffer.from(b)); },
        readDirectory: async (u) => fs.readdirSync(u.path, { withFileTypes: true }).map((d) => [d.name, d.isSymbolicLink() ? 64 | (fs.statSync(path.join(u.path, d.name)).isDirectory() ? 2 : 1) : d.isDirectory() ? 2 : 1]),
      },
      openTextDocument: async (u) => ({ uri: u }),
      onDidChangeConfiguration: () => ({ dispose() {} }),
    },
    commands: { executeCommand: async (...a) => { log.commands.push(a); }, registerCommand: () => ({ dispose() {} }) },
    env: { openExternal: async (u) => { log.external.push(String(u)); return true; } },
  };
  const context = { globalState: { get: (k) => state.get(k), update: async (k, v) => { state.set(k, v); } }, subscriptions: [] };
  return { vscode, context, log, answers, settings, Uri };
}

function fakeLicense(isPro) {
  const calls = { requirePro: [] };
  return { calls, pro: isPro, isPro() { return this.pro; }, refresh() { return Promise.resolve(); }, requirePro: async (f, detail) => { calls.requirePro.push({ f, detail }); return false; }, ready: Promise.resolve(), onDidChange: () => ({ dispose() {} }) };
}

// One open deck whose webview answers an export with fixture pictures (failing the slides in `bad`).
function setup(o) {
  o = o || {};
  const f = fakeVscode(o);
  const { createPro } = require('../src/pro');
  const { createExportTools } = require('../src/exportTools');
  const readDeck = READ_DECK;
  const lic = fakeLicense(!!o.pro);
  const pro = createPro({ vscode: f.vscode, context: f.context, license: lic });
  const slides = o.slides || 4;
  const bad = new Set(o.bad || []);
  let tools;
  const entry = {
    uri: f.Uri.file('/w/decks/Q3 review.pptx'), deck: { widthPx: 1280, heightPx: 720, slides: new Array(slides).fill({}) }, rendered: { slides },
    panel: { visible: true, webview: { postMessage: (m) => {
      if (m.type === 'exportSlides') {
        setImmediate(async () => {
          const failed = [];
          for (let n = 1; n <= slides; n++) {
            if (bad.has(n)) { failed.push(n); continue; }
            tools.onMessage(entry, parseWebviewMessage({ type: 'exportSlide', id: m.id, slide: n, data: new Uint8Array(m.format === 'jpeg' ? JPG : PNG) }));
            await new Promise((r) => setImmediate(r));
          }
          tools.onMessage(entry, parseWebviewMessage({ type: 'exportDone', id: m.id, failed, cancelled: false }));
        });
      }
      return Promise.resolve(true);
    } } },
  };
  const provider = { active: entry, entries: new Set([entry]) };
  tools = createExportTools({ vscode: f.vscode, provider, pro, readDeck });
  return Object.assign(f, { pro, lic, tools, entry });
}

// readDeck lives in extension.js, which needs the vscode module; load it once with a stub.
const vscodeStub = { workspace: {}, window: {}, commands: {}, Uri: {}, env: {} };
let READ_DECK;
{
  const orig = Module._load;
  Module._load = function (req, ...rest) { return req === 'vscode' ? vscodeStub : orig.call(this, req, ...rest); };
  READ_DECK = require('../src/extension').readDeck;
  Module._load = orig;
}

test('host: all slides as PNG into a folder, sorted names, one free try used', async () => {
  const t = setup({ open: [[{ path: '/w/out' }]] });
  t.answers.open[0] = [t.Uri.file('/w/out')];
  const out = await t.tools.exportPngs();
  assert.deepEqual(out.map((u) => u.path), ['/w/out/Q3 review-slide-01.png', '/w/out/Q3 review-slide-02.png', '/w/out/Q3 review-slide-03.png', '/w/out/Q3 review-slide-04.png']);
  assert.ok(out.every((u) => t.log.written.get(u.path).equals(PNG)));
  assert.equal(t.pro.left('images'), 2);
  assert.match(t.log.status[0], /free try of export all slides as png used, 2 left/);
  assert.match(t.log.info[0], /4 slides saved as PNG in out\./);
});

test('host: a slide that cannot be drawn is named in the message, the others are written', async () => {
  const t = setup({ bad: [2] });
  t.answers.open.push([t.Uri.file('/w/out')]);
  const out = await t.tools.exportPngs();
  assert.equal(out.length, 3);
  assert.match(t.log.info[0], /3 slides saved as PNG.*Slide 2 could not be drawn and is left out\./);
});

test('host: existing PNG files are only replaced after a yes; cancel uses no try', async () => {
  const t = setup();
  t.answers.open.push([t.Uri.file('/w/out')]);
  await t.tools.exportPngs();
  t.answers.open.push([t.Uri.file('/w/out')]);
  t.answers.warn.push(undefined);
  await t.tools.exportPngs();
  assert.ok(t.log.warn.some((w) => /4 of the 4 PNG files already exist/.test(w)));
  assert.equal(t.pro.left('images'), 2);
});

test('host: deck as PDF, one page per slide at the slide size', async () => {
  const t = setup({ save: ['default'], slides: 3 });
  const out = await t.tools.exportPdf();
  assert.equal(out.path, '/w/decks/Q3 review.pdf');
  const s = t.log.written.get(out.path).toString('latin1');
  assert.match(s, /\/Count 3 /);
  assert.equal((s.match(/\/MediaBox \[0 0 960 540\]/g) || []).length, 3);
  assert.match(s, /\/Title \(Q3 review\)/);
  assert.equal(t.pro.left('pdf'), 2);
  assert.equal(t.pro.left('images'), 3);
});

test('host: Escape in the dialogs uses no try and writes nothing', async () => {
  const t = setup({ save: [undefined], open: [undefined], pick: [undefined] });
  await t.tools.exportPdf();
  await t.tools.exportPngs();
  await t.tools.markdownBatch();
  assert.equal(t.log.written.size, 0);
  assert.equal(t.pro.left('pdf') + t.pro.left('images') + t.pro.left('markdownBatch'), 9);
});

test('host: after 3 tries the 4th explains Pro once and asks for nothing else', async () => {
  const t = setup({ save: ['default', 'default', 'default'] });
  for (let i = 0; i < 3; i++) await t.tools.exportPdf();
  const saves = t.log.saves.length;
  await t.tools.exportPdf();
  assert.equal(t.lic.calls.requirePro.length, 1);
  assert.equal(t.lic.calls.requirePro[0].f, 'Export the deck as PDF');
  assert.match(t.lic.calls.requirePro[0].detail, /used the 3 free tries/);
  assert.match(t.lic.calls.requirePro[0].detail, /USD 19 once/);
  assert.match(t.lic.calls.requirePro[0].detail, /stays free/);
  assert.equal(t.log.saves.length, saves);
});

test('host: a Pro user uses no tries', async () => {
  const t = setup({ pro: true, save: ['default'] });
  await t.tools.exportPdf();
  assert.equal(t.pro.left('pdf'), 3);
  assert.equal(t.log.status.length, 0);
});

test('host: text-only and old .ppt decks say why there is nothing to draw', async () => {
  const t = setup();
  t.entry.rendered = null;
  t.entry.uri = t.Uri.file('/w/old.ppt');
  await t.tools.exportPngs();
  assert.match(t.log.info[0], /shown as text only/);
  assert.equal(t.pro.left('images'), 3);
});

test('host: decks in a folder to Markdown, one file per deck next to it; broken decks are reported', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pro2-pptx-'));
  try {
    fs.mkdirSync(path.join(dir, 'sub'));
    fs.mkdirSync(path.join(dir, 'node_modules'));
    fs.copyFileSync(path.join(FIX, 'quarterly-review.pptx'), path.join(dir, 'quarterly-review.pptx'));
    fs.copyFileSync(path.join(FIX, 'legacy-basic.ppt'), path.join(dir, 'sub', 'legacy-basic.ppt'));
    fs.copyFileSync(path.join(FIX, 'legacy-broken.ppt'), path.join(dir, 'sub', 'legacy-broken.ppt'));
    fs.copyFileSync(path.join(FIX, 'quarterly-review.pptx'), path.join(dir, 'node_modules', 'skip.pptx'));
    fs.writeFileSync(path.join(dir, '~$lock.pptx'), 'lock');
    const t = setup({ pick: ['each'] });
    t.provider = null;
    const out = await t.tools.markdownBatch(t.Uri.file(dir));
    const names = out.map((u) => path.relative(dir, u.path)).sort();
    assert.deepEqual(names, ['quarterly-review.md', path.join('sub', 'legacy-basic.md')]);
    const md = t.log.written.get(path.join(dir, 'quarterly-review.md')).toString('utf8');
    assert.match(md, /^# quarterly-review\.pptx\n\n## /);
    assert.match(md, /> \*\*Notes:\*\*/);
    assert.match(t.log.info[0], /2 decks converted to Markdown\. 1 could not be read: sub\/legacy-broken\.ppt\./);
    assert.equal(t.pro.left('markdownBatch'), 2);
    // One file for all decks.
    t.answers.pick.push('one'); t.answers.save.push('default');
    const one = await t.tools.markdownBatch(t.Uri.file(dir));
    const all = t.log.written.get(one[0].path).toString('utf8');
    assert.equal(path.basename(one[0].path), 'slides.md');
    assert.match(all, /^# Slides\n\n2 decks, \d+ slides\./);
    assert.match(all, /## quarterly-review\.pptx\n\n### /);
    assert.match(all, /## Not converted\n\n- sub\/legacy-broken\.ppt: /);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('host: a folder with no decks says so and uses no try', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pro2-pptx-empty-'));
  try {
    const t = setup();
    await t.tools.markdownBatch(t.Uri.file(dir));
    assert.match(t.log.info[0], /No PowerPoint decks/);
    assert.equal(t.pro.left('markdownBatch'), 3);
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('host: showProFeatures=false hides the Export button unless Pro', () => {
  const { createPro } = require('../src/pro');
  const f = fakeVscode({ settings: { showProFeatures: false } });
  const lic = fakeLicense(false);
  const pro = createPro({ vscode: f.vscode, context: f.context, license: lic });
  assert.equal(pro.shown(), false);
  lic.pro = true;
  assert.equal(pro.shown(), true);
  assert.equal(pro.label('pdf'), '');
});

// ---------- activation ----------
test('activate: no message and no network at start; every Pro command is registered and declared', async () => {
  const f = fakeVscode();
  const registered = [];
  f.vscode.commands.registerCommand = (id) => { registered.push(id); return { dispose() {} }; };
  f.vscode.window.registerCustomEditorProvider = () => ({ dispose() {} });
  f.vscode.window.registerUriHandler = () => ({ dispose() {} });
  f.vscode.env.uriScheme = 'vscode'; f.vscode.env.appName = 'Visual Studio Code';
  f.vscode.authentication = { getSession: async () => undefined, onDidChangeSessions: () => ({ dispose() {} }) };
  f.context.secrets = { get: async () => undefined, store: async () => {}, delete: async () => {}, onDidChange: () => ({ dispose() {} }) };
  f.context.globalState.setKeysForSync = () => {};
  f.context.extensionUri = f.Uri.file('/ext');
  const realFetch = globalThis.fetch;
  let fetched = 0;
  globalThis.fetch = async () => { fetched++; throw new Error('no network in tests'); };
  const orig = Module._load;
  Module._load = function (req, ...rest) { return req === 'vscode' ? f.vscode : orig.call(this, req, ...rest); };
  try {
    delete require.cache[require.resolve('../src/extension')];
    const api = require('../src/extension').activate(f.context);
    await api.pro.ready;
    await new Promise((r) => setTimeout(r, 20));
    for (const id of ['pptxViewer.exportMenu', 'pptxViewer.exportSlidesPng', 'pptxViewer.exportPdf', 'pptxViewer.decksToMarkdown', 'pptxViewer.getPro', 'pptxViewer.unlockWithSponsors', 'pptxViewer.enterLicenseKey', 'pptxViewer.removeLicense']) assert.ok(registered.includes(id), id);
    const declared = require('../package.json').contributes.commands.map((c) => c.command);
    for (const id of registered) assert.ok(declared.includes(id), 'declared: ' + id);
    assert.equal(f.log.info.length + f.log.warn.length + f.log.modal.length, 0);
    assert.equal(fetched, 0);
    assert.deepEqual(api.provider.proState(), { show: true, badge: true });
  } finally {
    Module._load = orig;
    globalThis.fetch = realFetch;
  }
});

test('package.json: Pro commands hide with the setting; free commands keep their old rules', () => {
  const pkg = require('../package.json');
  const pal = pkg.contributes.menus.commandPalette;
  for (const id of ['pptxViewer.exportMenu', 'pptxViewer.exportPdf', 'pptxViewer.decksToMarkdown', 'pptxViewer.getPro']) assert.match(pal.find((m) => m.command === id).when, /config\.pptxViewer\.showProFeatures/, id);
  for (const m of pal) if (!/export|decksToMarkdown|Pro|License|Sponsors/.test(m.command)) assert.doesNotMatch(m.when || '', /showProFeatures/, m.command);
  assert.equal(pkg.contributes.configuration.properties['pptxViewer.showProFeatures'].default, true);
  assert.ok(pkg.contributes.menus['explorer/context'].some((m) => m.command === 'pptxViewer.decksToMarkdown' && /explorerResourceIsFolder/.test(m.when)));
});
