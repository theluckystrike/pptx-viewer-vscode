'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { parseWebviewMessage, readConfig } = require('../src/protocol');
const { buildHtml, buildCsp } = require('../src/html');

test('accepts known messages with valid values', () => {
  assert.deepEqual(parseWebviewMessage({ type: 'ready' }), { type: 'ready' });
  assert.deepEqual(parseWebviewMessage({ type: 'copySlideMarkdown', slide: 3 }), { type: 'copySlideMarkdown', slide: 3 });
  assert.deepEqual(parseWebviewMessage({ type: 'copyDeckMarkdown', extra: 1 }), { type: 'copyDeckMarkdown' });
  assert.deepEqual(parseWebviewMessage({ type: 'rendered', slides: 6, ms: 45 }), { type: 'rendered', slides: 6, ms: 45 });
  assert.deepEqual(parseWebviewMessage({ type: 'slideChanged', slide: 2 }), { type: 'slideChanged', slide: 2 });
});

test('rejects unknown types and bad values', () => {
  for (const m of [null, 'ready', {}, { type: 'openExternal', url: 'https://x' }, { type: 'copySlideMarkdown', slide: 0 },
    { type: 'copySlideMarkdown', slide: '2' }, { type: 'copySlideMarkdown', slide: 1.5 }, { type: 'slideChanged', slide: -1 },
    { type: 'rendered', slides: 'x' }, { type: '__proto__' }]) {
    assert.equal(parseWebviewMessage(m), null, JSON.stringify(m));
  }
});

test('error messages are truncated', () => {
  assert.equal(parseWebviewMessage({ type: 'error', message: 'x'.repeat(2000) }).message.length, 500);
  assert.equal(parseWebviewMessage({ type: 'error', message: 5 }).message, 'Unknown error');
});

test('config defaults and size limit', () => {
  const c = readConfig((k, d) => d);
  assert.equal(c.maxFileSizeBytes, 150 * 1024 * 1024);
  assert.equal(c.showNotes, true);
  assert.equal(c.renderSlides, true);
  assert.equal(readConfig((k, d) => (k === 'maxFileSizeMB' ? -5 : d)).maxFileSizeBytes, 150 * 1024 * 1024);
  assert.equal(readConfig((k, d) => (k === 'renderSlides' ? false : d)).renderSlides, false);
});

test('CSP: nonce scripts only, no network, no eval', () => {
  const csp = buildCsp('vscode-resource:', 'abc');
  assert.match(csp, /default-src 'none'/);
  assert.match(csp, /script-src 'nonce-abc'(;|$)/);
  assert.match(csp, /connect-src 'none'/);
  assert.ok(!csp.includes('unsafe-eval'));
  assert.ok(!/https?:/.test(csp));
  const html = buildHtml({ css: 'c"<', core: 'a', viewer: 'b', preview: 'p' }, 'vscode-resource:', 'abc');
  assert.equal((html.match(/<script /g) || []).length, 3);
  assert.equal((html.match(/nonce="abc"/g) || []).length, 3);
  assert.ok(html.includes('href="c&quot;&lt;"'));
});
