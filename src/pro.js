'use strict';
// PPTX Viewer Pro on top of the shared Zovo license module (src/license, copied unchanged from
// factory/sprint/license). Pro features: export every slide as PNG files, export the deck as one PDF, and convert
// a folder of decks to Markdown with speaker notes. Everything that was free stays free.
// Rules: nothing about Pro shows on open; each Pro feature works TRIAL_USES times on this computer first (a try is
// used only when a file was really written); after that a click explains it once (a modal the user started) with
// the ways to unlock. No network except what the user starts (the GitHub Sponsors unlock, or a page in the browser).
const { createLicense } = require('./license');

const PRO_ENABLED = true;
const SKU = 'pptx-pro';
const TRIAL_USES = 3;
const TRIAL_KEY = 'pptxViewer.proTrial';

const FEATURES = {
  images: {
    title: 'Export all slides as PNG',
    detail: 'Saves every slide of the deck as a PNG picture (about 1920 pixels wide) in a folder you pick, named so they sort in slide order. Good for docs, wikis, tickets and release notes.'
  },
  pdf: {
    title: 'Export the deck as PDF',
    detail: 'Saves the whole deck as one PDF, one slide per page at the slide size, made on your computer without PowerPoint or LibreOffice. Good for sharing a deck with someone who has no PowerPoint.'
  },
  markdownBatch: {
    title: 'Decks in a folder to Markdown',
    detail: 'Converts every PowerPoint deck in a folder and its subfolders (.pptx, .ppsx, .potx, .pptm and old .ppt) to Markdown with titles, bullet levels, tables and speaker notes: one .md file per deck, or one file for all of them. Good for search, docs and AI tools.'
  }
};

const PRICE_LINE = 'Zovo Pro is USD 19 once and unlocks the Pro features in every Zovo viewer. Already bought the "$19 one time" tier on GitHub Sponsors? Choose Unlock with GitHub Sponsors: it asks GitHub one question and sends nothing else. Paid with PayPal? Choose I have a license key.';
const FREE_LINE = 'Everything else in PPTX Viewer stays free: viewing, Present, notes, search, thumbnails, the outline, copy or save one slide as PNG, and Copy as Markdown for a slide or the whole deck.';

function createPro({ vscode, context, license, enabled }) {
  const on = enabled === undefined ? PRO_ENABLED : enabled;
  const lic = license || createLicense({ vscode, context, sku: SKU, name: 'Zovo', commandPrefix: 'pptxViewer' });
  const used = (f) => Number((context.globalState.get(TRIAL_KEY) || {})[f]) || 0;
  const left = (f) => (FEATURES[f] ? Math.max(0, TRIAL_USES - used(f)) : 0);
  const pro = () => !on || lic.isPro(SKU);

  async function explain(f) {
    const ft = FEATURES[f];
    const tried = used(f) >= TRIAL_USES ? 'You have used the ' + TRIAL_USES + ' free tries of this feature on this computer.\n\n' : '';
    return lic.requirePro(ft.title, tried + ft.detail + '\n\n' + PRICE_LINE + '\n\n' + FREE_LINE);
  }

  // Call when the user has chosen a Pro action, before asking for its inputs. True when it may run: Pro, or free
  // tries left. With no tries left it shows the one explaining modal (user-started) and returns the unlock result.
  async function check(f) {
    if (pro()) { if (on) lic.refresh(); return true; }
    if (left(f) > 0) return true;
    return explain(f);
  }

  // Call after the action wrote its output. Uses one free try when the user is not Pro.
  async function commit(f) {
    if (pro() || !FEATURES[f]) return;
    const m = Object.assign({}, context.globalState.get(TRIAL_KEY) || {});
    m[f] = used(f) + 1;
    await context.globalState.update(TRIAL_KEY, m);
    const n = left(f);
    vscode.window.setStatusBarMessage('PPTX Viewer: free try of ' + FEATURES[f].title.toLowerCase() + ' used, ' + n + ' left on this computer.', 6000);
  }

  // pptxViewer.showProFeatures = false hides the Export button and the Pro commands (the "do not show again" for Pro).
  const shown = () => pro() || vscode.workspace.getConfiguration('pptxViewer').get('showProFeatures', true) !== false;

  return {
    enabled: on,
    shown,
    ready: lic.ready,
    check,
    commit,
    left,
    isPro: pro,
    label: (f) => (pro() ? '' : left(f) > 0 ? 'Pro, ' + left(f) + ' free ' + (left(f) === 1 ? 'try' : 'tries') : 'Pro'),
    onDidChange: (fn) => lic.onDidChange(fn),
    license: lic
  };
}

module.exports = { createPro, FEATURES, SKU, TRIAL_USES, TRIAL_KEY, PRO_ENABLED, PRICE_LINE };
