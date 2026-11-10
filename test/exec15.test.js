'use strict';
// EXEC15: PPTX 0.5.1. CONV3 fix 6 (one "Zovo Pro" palette category with the same 5 titles in every Zovo extension)
// and the Sponsors-first Pro window from the shared license module.
const test = require('node:test');
const assert = require('node:assert/strict');
const { createPro, TRIAL_USES } = require('../src/pro');

test('EXEC15: the 5 Zovo Pro commands share the "Zovo Pro" category, with the same titles as every Zovo extension', () => {
  const pk = require('../package.json');
  const zovo = pk.contributes.commands.filter((c) => c.category === 'Zovo Pro');
  assert.deepEqual(zovo.map((c) => [c.command, c.title]), [
    ['pptxViewer.getPro', 'Get Zovo Pro (USD 19 once, all 22 Zovo extensions)'],
    ['pptxViewer.payWithPayPal', 'Pay for Zovo Pro with PayPal (USD 19, team of 5 USD 59)'],
    ['pptxViewer.unlockWithSponsors', 'I already paid: unlock with GitHub Sponsors'],
    ['pptxViewer.enterLicenseKey', 'Enter Zovo Pro license key'],
    ['pptxViewer.removeLicense', 'Remove Zovo Pro from this editor']
  ]);
  // The Pro tools keep "(Pro)" in their own titles, so typing "Pro" in the palette finds them.
  for (const id of ['exportMenu', 'exportSlidesPng', 'exportPdf', 'decksToMarkdown']) assert.match(pk.contributes.commands.find((c) => c.command === 'pptxViewer.' + id).title, /\(Pro\)/);
  assert.equal(pk.engines.vscode, '^1.90.0');
});

test('EXEC15: the Pro window (module v2) has 4 buttons, Get Zovo Pro first, and only shows after the 3 free tries were used', async () => {
  const modal = [];
  const store = new Map();
  const vscode = {
    Uri: { parse: (s) => ({ toString: () => s, s }) },
    env: { openExternal: async () => true },
    window: {
      setStatusBarMessage() {},
      showInformationMessage: async (msg, o, ...b) => { if (o && o.modal) modal.push({ msg, detail: o.detail, b }); return undefined; },
      showWarningMessage: async () => undefined,
      showInputBox: async () => undefined,
      showQuickPick: async () => undefined
    },
    workspace: { getConfiguration: () => ({ get: (k, d) => d }) },
    commands: { registerCommand: () => ({ dispose() {} }) }
  };
  const ctx = { globalState: { get: (k) => store.get(k), update: async (k, v) => { store.set(k, v); } }, subscriptions: [], secrets: { get: async () => undefined, store: async () => {}, delete: async () => {}, onDidChange: () => ({ dispose() {} }) } };
  const pro = createPro({ vscode, context: ctx });
  await pro.ready;
  for (let i = 0; i < TRIAL_USES; i++) { assert.equal(await pro.check('pdf'), true); await pro.commit('pdf'); }
  assert.equal(modal.length, 0, 'no Pro window while tries are left');
  assert.equal(await pro.check('pdf'), false);
  assert.equal(modal.length, 1);
  assert.deepEqual(modal[0].b, ['Get Zovo Pro, USD 19', 'Pay with PayPal', 'I already paid']);
  assert.match(modal[0].msg, /^Export the deck as PDF is part of Zovo Pro\./);
  assert.match(modal[0].detail, /It turns on Export all slides as PNG and Decks in a folder to Markdown here, and the Pro tools of all 22 Zovo extensions/);
  assert.match(modal[0].detail, /Everything that is free today stays free\./);
  assert.doesNotMatch(modal[0].detail, /refund/i);
  assert.match(modal[0].detail, /USD 19/);
});
