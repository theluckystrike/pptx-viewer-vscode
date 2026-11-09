'use strict';
// Zovo Pro exports (extension host side): every slide as PNG, the deck as one PDF, and a folder of decks to
// Markdown. Slide pictures are drawn by the webview (the same code as the free "Save as PNG") and sent here one
// at a time; this file asks where to save, checks every picture and writes the files. No network.
const path = require('path');
const { imagesToPdf, slideSizePt, jpegInfo } = require('./pdfwriter');
const { deckToMarkdown, mdEscapeInline } = require('./pptx');
const { isPng } = require('./protocol');

const SLIDE_TIMEOUT_MS = 60000; // per slide; a slow slide resets the clock
const DECK_EXT = /\.(pptx|ppsx|potx|pptm|ppt|pps|pot)$/i;
const SKIP_DIRS = new Set(['node_modules', '.git', '.svn', '.hg', '__pycache__', '.venv', 'venv']);
const MAX_DECKS = 2000;
const MAX_DEPTH = 12;
const MAX_DECK_BYTES = 300 * 1024 * 1024;

/** "deck-slide-07.png": zero-padded to the slide count so the files sort in slide order. */
function slideFileName(base, n, total, ext) {
  const w = Math.max(2, String(total).length);
  return base + '-slide-' + String(n).padStart(w, '0') + '.' + ext;
}

/** Deck file name without its extension. */
function deckBase(fileName) {
  return String(fileName || 'deck').replace(/\.[^.\/]+$/, '') || 'deck';
}

/** One Markdown file for many decks: a heading per deck (its path), slides one level below. */
function combinedMarkdown(items) {
  const lines = ['# Slides', ''];
  const ok = items.filter((x) => x.markdown);
  lines.push(ok.length + ' deck' + (ok.length === 1 ? '' : 's') + ', ' + ok.reduce((n, x) => n + x.slides, 0) + ' slides.', '');
  for (const x of ok) {
    // deckToMarkdown gives "# file" then "## slide" headings: one level down under the deck's path.
    const body = x.markdown.replace(/^# .*\n/, '').replace(/^## /gm, '### ').trim();
    lines.push('## ' + mdEscapeInline(x.rel), '', body, '');
  }
  const bad = items.filter((x) => !x.markdown);
  if (bad.length) {
    lines.push('## Not converted', '');
    for (const x of bad) lines.push('- ' + mdEscapeInline(x.rel) + ': ' + mdEscapeInline(x.error || 'unknown error'));
    lines.push('');
  }
  return lines.join('\n');
}

function createExportTools({ vscode, provider, pro, readDeck }) {
  let seq = 0;
  const jobs = new Map(); // id -> job

  const say = (m, ...b) => vscode.window.showInformationMessage(m, ...b);
  const warn = (m) => vscode.window.showWarningMessage(m);
  const nameOf = (u) => path.posix.basename(u.path);
  const dirOf = (u) => vscode.Uri.joinPath(u, '..');

  function activeDeck() {
    const e = provider.active || [...(provider.entries || [])].find((x) => x.panel && x.panel.visible);
    if (!e || !e.deck) { say('Open a PowerPoint deck in PPTX Viewer first.'); return null; }
    if (!e.rendered || !e.rendered.slides) {
      say(e.deck.textOnly || /\.(ppt|pps|pot)$/i.test(e.uri.path) ? 'This deck is shown as text only, so there are no slide pictures to export. "Decks in a folder to Markdown" works for it.' : 'The slides are still being drawn. Try again in a moment.');
      return null;
    }
    return e;
  }

  // Asks the webview for every slide picture. onSlide(n, bytes) is awaited per slide. Resolves
  // { count, failed: [n], cancelled } or { error }.
  function runExport(entry, format, onSlide, progress, token) {
    const id = ++seq;
    return new Promise((resolve) => {
      const job = { id, entry, format, onSlide, progress, resolve, timer: null, chain: Promise.resolve(), done: false, total: entry.rendered.slides, got: 0 };
      const fail = (error) => { if (job.done) return; job.done = true; clearTimeout(job.timer); jobs.delete(id); resolve({ error }); };
      job.arm = () => { clearTimeout(job.timer); job.timer = setTimeout(() => fail('A slide took more than a minute to draw.'), SLIDE_TIMEOUT_MS); };
      job.fail = fail;
      jobs.set(id, job);
      job.arm();
      if (token) token.onCancellationRequested(() => { entry.panel.webview.postMessage({ type: 'exportCancel', id }); });
      entry.panel.webview.postMessage({ type: 'exportSlides', id, format });
    });
  }

  // From the provider: 'exportSlide' and 'exportDone' (already checked by protocol.js).
  function onMessage(entry, m) {
    const job = jobs.get(m.id);
    if (!job || job.entry !== entry || job.done) return;
    if (m.type === 'exportSlide') {
      job.arm();
      const okType = job.format === 'png' ? isPng(m.data) : !!jpegInfo(m.data);
      if (!okType) return job.fail('Slide ' + m.slide + ' did not come back as a picture.');
      job.got++;
      if (job.progress) job.progress.report({ message: 'slide ' + m.slide + ' of ' + job.total, increment: 100 / Math.max(1, job.total) });
      job.chain = job.chain.then(() => job.onSlide(m.slide, m.data)).catch((e) => job.fail(String((e && e.message) || e)));
    } else if (m.type === 'exportDone') {
      job.chain.then(() => {
        if (job.done) return;
        job.done = true;
        clearTimeout(job.timer);
        jobs.delete(job.id);
        job.resolve(m.error ? { error: m.error } : { count: job.got, failed: m.failed || [], cancelled: !!m.cancelled });
      });
    }
  }

  function onPanelClosed(entry) {
    for (const job of [...jobs.values()]) if (job.entry === entry) job.fail('The deck was closed.');
  }

  const progressOpts = (title) => ({ location: vscode.ProgressLocation.Notification, title, cancellable: true });
  const failedNote = (failed) => (failed && failed.length ? ' Slide' + (failed.length === 1 ? ' ' : 's ') + failed.join(', ') + (failed.length === 1 ? ' could' : ' could') + ' not be drawn and ' + (failed.length === 1 ? 'is' : 'are') + ' left out.' : '');

  async function exportPngs() {
    const entry = activeDeck();
    if (!entry || !(await pro.check('images'))) return;
    const base = deckBase(nameOf(entry.uri));
    const picked = await vscode.window.showOpenDialog({ title: 'Folder for the ' + entry.rendered.slides + ' slide pictures', canSelectFiles: false, canSelectFolders: true, canSelectMany: false, defaultUri: dirOf(entry.uri), openLabel: 'Save here' });
    if (!picked || !picked[0]) return;
    const folder = picked[0];
    const total = entry.rendered.slides;
    const targets = Array.from({ length: total }, (_, i) => vscode.Uri.joinPath(folder, slideFileName(base, i + 1, total, 'png')));
    let existing = 0;
    for (const t of targets) { try { await vscode.workspace.fs.stat(t); existing++; } catch { /* free */ } }
    if (existing) {
      const ok = await vscode.window.showWarningMessage(existing + ' of the ' + total + ' PNG files already exist in that folder.', { modal: true }, 'Replace them');
      if (ok !== 'Replace them') return;
    }
    const written = [];
    const r = await vscode.window.withProgress(progressOpts('Exporting ' + total + ' slides as PNG'), (progress, token) =>
      runExport(entry, 'png', async (n, bytes) => { const t = targets[n - 1]; await vscode.workspace.fs.writeFile(t, bytes); written.push(t); }, progress, token));
    if (r.error) return warn('Could not export the slides: ' + r.error + (written.length ? ' ' + written.length + ' PNG files were written.' : ''));
    if (r.cancelled) return say('Export stopped. ' + written.length + ' PNG file' + (written.length === 1 ? ' was' : 's were') + ' written.');
    if (!written.length) return warn('No slide could be drawn, so no PNG file was written.');
    await pro.commit('images');
    const pick = await say(written.length + ' slides saved as PNG in ' + nameOf(folder) + '.' + failedNote(r.failed), 'Reveal in Explorer');
    if (pick === 'Reveal in Explorer') await vscode.commands.executeCommand('revealInExplorer', written[0]);
    return written;
  }

  async function exportPdf() {
    const entry = activeDeck();
    if (!entry || !(await pro.check('pdf'))) return;
    const base = deckBase(nameOf(entry.uri));
    const target = await vscode.window.showSaveDialog({ title: 'Save the deck as PDF', defaultUri: vscode.Uri.joinPath(dirOf(entry.uri), base + '.pdf'), filters: { PDF: ['pdf'] }, saveLabel: 'Export PDF' });
    if (!target) return;
    const pages = [];
    const r = await vscode.window.withProgress(progressOpts('Exporting ' + entry.rendered.slides + ' slides to PDF'), (progress, token) =>
      runExport(entry, 'jpeg', async (n, bytes) => { pages.push({ n, jpeg: bytes }); }, progress, token));
    if (r.error) return warn('Could not export the PDF: ' + r.error);
    if (r.cancelled) return say('Export stopped. No PDF was written.');
    if (!pages.length) return warn('No slide could be drawn, so no PDF was written.');
    pages.sort((a, b) => a.n - b.n);
    const pdf = imagesToPdf(pages, Object.assign({ title: base }, slideSizePt(entry.deck)));
    await vscode.workspace.fs.writeFile(target, pdf);
    await pro.commit('pdf');
    const pick = await say(pages.length + ' slides saved as ' + nameOf(target) + '.' + failedNote(r.failed), 'Open');
    if (pick === 'Open') await vscode.env.openExternal(target);
    return target;
  }

  // Every deck under a folder, depth-first, sorted by path. Skips dependency and VCS folders.
  async function findDecks(root) {
    const out = [];
    const walk = async (dir, depth) => {
      if (depth > MAX_DEPTH || out.length >= MAX_DECKS) return;
      let list;
      try { list = await vscode.workspace.fs.readDirectory(dir); } catch { return; }
      list.sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }));
      for (const [name, type] of list) {
        if (out.length >= MAX_DECKS) return;
        const u = vscode.Uri.joinPath(dir, name);
        if (type & 2) { if (!SKIP_DIRS.has(name) && !(type & 64)) await walk(u, depth + 1); } // 2 = Directory, 64 = SymbolicLink
        else if (DECK_EXT.test(name) && !name.startsWith('~$')) out.push(u);
      }
    };
    await walk(root, 0);
    return out;
  }

  async function markdownBatch(folderArg) {
    if (!(await pro.check('markdownBatch'))) return;
    let root = folderArg && folderArg.path ? folderArg : null;
    if (!root) {
      const ws = (vscode.workspace.workspaceFolders || [])[0];
      const picked = await vscode.window.showOpenDialog({ title: 'Folder with PowerPoint decks', canSelectFiles: false, canSelectFolders: true, canSelectMany: false, defaultUri: provider.active ? dirOf(provider.active.uri) : ws && ws.uri, openLabel: 'Convert decks' });
      if (!picked || !picked[0]) return;
      root = picked[0];
    }
    const decks = await findDecks(root);
    if (!decks.length) return say('No PowerPoint decks (.pptx, .ppsx, .potx, .pptm, .ppt) in ' + nameOf(root) + ' or its subfolders.');
    const mode = await vscode.window.showQuickPick([
      { label: 'One Markdown file per deck', description: 'next to each deck, ' + decks.length + ' files', id: 'each' },
      { label: 'One Markdown file for all decks', description: decks.length + ' decks in one file', id: 'one' },
    ], { title: 'Convert ' + decks.length + ' deck' + (decks.length === 1 ? '' : 's') + ' to Markdown with speaker notes (Zovo Pro)' });
    if (!mode) return;
    let oneTarget = null;
    if (mode.id === 'one') {
      oneTarget = await vscode.window.showSaveDialog({ title: 'Save all decks as one Markdown file', defaultUri: vscode.Uri.joinPath(root, 'slides.md'), filters: { Markdown: ['md'] } });
      if (!oneTarget) return;
    } else {
      let existing = 0;
      for (const d of decks) { try { await vscode.workspace.fs.stat(d.with({ path: d.path.replace(DECK_EXT, '.md') })); existing++; } catch { /* free */ } }
      if (existing) {
        const ok = await vscode.window.showWarningMessage(existing + ' of the ' + decks.length + ' Markdown files already exist next to their decks.', { modal: true }, 'Replace them');
        if (ok !== 'Replace them') return;
      }
    }
    const items = [];
    const written = [];
    await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: 'Converting ' + decks.length + ' decks to Markdown', cancellable: true }, async (progress, token) => {
      for (const d of decks) {
        if (token && token.isCancellationRequested) break;
        const rel = path.posix.relative(root.path, d.path) || nameOf(d);
        progress.report({ message: rel, increment: 100 / decks.length });
        try {
          const st = await vscode.workspace.fs.stat(d);
          if (st.size > MAX_DECK_BYTES) throw new Error('larger than 300 MB');
          const data = await vscode.workspace.fs.readFile(d);
          const { deck } = readDeck(data, nameOf(d));
          const markdown = deckToMarkdown(deck, nameOf(d));
          items.push({ rel, markdown, slides: deck.slides.length });
          if (mode.id === 'each') {
            const t = d.with({ path: d.path.replace(DECK_EXT, '.md') });
            await vscode.workspace.fs.writeFile(t, Buffer.from(markdown, 'utf8'));
            written.push(t);
          }
        } catch (e) {
          items.push({ rel, error: String((e && e.message) || e).slice(0, 200) });
        }
      }
    });
    const ok = items.filter((x) => x.markdown);
    if (!ok.length) return warn('No deck could be converted. ' + (items[0] ? items[0].rel + ': ' + items[0].error : ''));
    if (oneTarget) {
      await vscode.workspace.fs.writeFile(oneTarget, Buffer.from(combinedMarkdown(items), 'utf8'));
      written.push(oneTarget);
    }
    await pro.commit('markdownBatch');
    const bad = items.length - ok.length;
    const msg = ok.length + ' deck' + (ok.length === 1 ? '' : 's') + ' converted to Markdown' + (oneTarget ? ' in ' + nameOf(oneTarget) : '') + '.' + (bad ? ' ' + bad + ' could not be read' + (oneTarget ? ' (listed at the end of the file).' : ': ' + items.filter((x) => !x.markdown).map((x) => x.rel).slice(0, 5).join(', ') + '.') : '');
    const pick = await say(msg, 'Open');
    if (pick === 'Open') await vscode.window.showTextDocument(await vscode.workspace.openTextDocument(written[0]));
    return written;
  }

  // The Export button and "PPTX: Export (Pro)": one quick pick with the three Pro exports.
  async function menu() {
    const tag = (f) => { const l = pro.label(f); return l ? l : ''; };
    const pick = await vscode.window.showQuickPick([
      { label: 'Export all slides as PNG', description: tag('images'), id: 'png' },
      { label: 'Export the deck as PDF', description: tag('pdf'), id: 'pdf' },
      { label: 'Convert decks in a folder to Markdown with notes', description: tag('markdownBatch'), id: 'md' },
    ], { title: 'Export (Zovo Pro)', placeHolder: pro.isPro() ? 'Pick an export' : 'Each export has 3 free tries on this computer. Zovo Pro is USD 19 once, for every Zovo viewer.' });
    if (!pick) return;
    return { png: exportPngs, pdf: exportPdf, md: () => markdownBatch() }[pick.id]();
  }

  return { exportPngs, exportPdf, markdownBatch, menu, onMessage, onPanelClosed, findDecks, _jobs: jobs };
}

module.exports = { createExportTools, slideFileName, deckBase, combinedMarkdown, SLIDE_TIMEOUT_MS, DECK_EXT };
