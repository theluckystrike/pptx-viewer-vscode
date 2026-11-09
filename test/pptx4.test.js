'use strict';
// 0.4.0: renderer patches (master text styles, anchors, bullets, fonts), SmartArt notice only when
// there is no drawn copy, empty files, the local file poll, and the toolbar/undrawable texts.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const { zipSync, strToU8 } = require('fflate');
const { patchRenderer, PATCHES, RENDERER_VERSION, HELPERS } = require('../scripts/patch-renderer');
const { parseDeck } = require('../src/pptx');
const core = require('../media/core');

const ROOT = path.join(__dirname, '..');
const BUNDLE = path.join(ROOT, 'node_modules/pptx-preview/dist/pptx-preview.umd.js');

test('every renderer patch matches the pinned pptx-preview exactly as often as expected', () => {
  assert.equal(require(path.join(ROOT, 'node_modules/pptx-preview/package.json')).version, RENDERER_VERSION);
  assert.equal(require(path.join(ROOT, 'package.json')).dependencies['pptx-preview'], RENDERER_VERSION);
  const src = fs.readFileSync(BUNDLE, 'utf8');
  const { src: out, applied } = patchRenderer(src);
  assert.deepEqual(applied, PATCHES.map((p) => p.id));
  for (const p of PATCHES) assert.equal((out.match(p.find) || []).length, 0, p.id + ' left an unpatched copy');
  // The patched bundle still parses as JavaScript.
  assert.doesNotThrow(() => new vm.Script(out));
  assert.match(out, /"titleStyle":\["dt","ftr","sldNum","hdr"\]\.includes\(i\)\?"otherStyle":"bodyStyle"/);
});

test('a patch that no longer matches stops the build', () => {
  assert.throws(() => patchRenderer('var nothing = 1;'), /matched 0 times/);
});

test('the shipped renderer is the patched one', () => {
  const vendor = path.join(ROOT, 'media/vendor/pptx-preview.umd.js');
  if (!fs.existsSync(vendor)) return; // not built yet
  const text = fs.readFileSync(vendor, 'utf8');
  for (const p of PATCHES) assert.ok(text.includes(p.id), p.id);
});

function helpers() {
  const ctx = {};
  vm.createContext(ctx);
  vm.runInContext(HELPERS, ctx);
  return ctx;
}

test('fonts: a:latin first, theme fonts resolved, a:ea only as the last choice', () => {
  const h = helpers();
  const theme = { source: { 'a:theme': { 'a:themeElements': { 'a:fontScheme': {
    'a:majorFont': { 'a:latin': { attrs: { typeface: 'Aptos Display' } }, 'a:ea': { attrs: { typeface: '' } } },
    'a:minorFont': { 'a:latin': { attrs: { typeface: 'Aptos' } } },
  } } } } };
  const rpr = (latin, ea) => Object.assign({}, latin && { 'a:latin': { attrs: { typeface: latin } } }, ea && { 'a:ea': { attrs: { typeface: ea } } });
  assert.equal(h.__zovoFace(rpr('+mn-lt'), theme), 'Aptos');
  assert.equal(h.__zovoFace(rpr('+mj-lt', '+mj-ea'), theme), 'Aptos Display');
  assert.equal(h.__zovoFace(rpr(null, '+mj-ea'), theme), 'Aptos Display'); // empty ea falls back to latin
  assert.equal(h.__zovoFace(rpr('Georgia', 'DengXian'), theme), 'Georgia');
  assert.equal(h.__zovoFace(rpr(null, 'DengXian'), theme), 'DengXian');
  assert.equal(h.__zovoFace({}, theme), undefined);
  assert.equal(h.__zovoFace(rpr('+mn-lt'), undefined), undefined);
});

test('fonts: every family gets fallbacks and a generic family; names cannot break out of CSS', () => {
  const h = helpers();
  assert.equal(h.__zovoStack('Calibri'), '"Calibri",Carlito,"Segoe UI","Helvetica Neue",Arial,sans-serif');
  assert.match(h.__zovoStack('Cambria'), /Caladea.*serif$/);
  assert.match(h.__zovoStack('Courier New'), /monospace$/);
  assert.match(h.__zovoStack('Some Brand Font'), /^"Some Brand Font",sans-serif$/);
  assert.equal(h.__zovoStack('x";}body{color:red'), '"xbodycolor:red",sans-serif');
});

function deckWithSmartArt(withDrawing) {
  const slideRels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>',
    withDrawing ? '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>' : '',
  ].join('');
  const files = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>',
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>',
    'ppt/slides/_rels/slide1.xml.rels': `<Relationships>${slideRels}</Relationships>`,
    'ppt/diagrams/data1.xml': '<dgm:dataModel xmlns:dgm="d"/>',
  };
  if (withDrawing) files['ppt/diagrams/drawing1.xml'] = '<dsp:drawing xmlns:dsp="s"/>';
  const z = {};
  for (const [k, v] of Object.entries(files)) z[k] = strToU8(v);
  return zipSync(z);
}

test('SmartArt saved by PowerPoint (with its drawn copy) gets no "cannot draw" notice', () => {
  assert.deepEqual(parseDeck(deckWithSmartArt(true)).slides[0].notices, []);
  assert.deepEqual(parseDeck(deckWithSmartArt(false)).slides[0].notices, ['smartart']);
  assert.equal('diagrams' in parseDeck(deckWithSmartArt(true)).slides[0], false);
});

test('webview texts: a slide that cannot be drawn says so', () => {
  assert.match(core.UNDRAWABLE_MESSAGE, /could not be drawn/);
  assert.match(core.UNDRAWABLE_MESSAGE, /PowerPoint, Keynote or LibreOffice/);
});

test('the webview greys out the toolbar until a deck with slides is open', () => {
  const js = fs.readFileSync(path.join(ROOT, 'media/viewer.js'), 'utf8');
  const ids = js.match(/DECK_CONTROLS = \[([^\]]+)\]/)[1];
  for (const id of ['prevBtn', 'nextBtn', 'slideInput', 'presentBtn', 'searchInput', 'mdBtn', 'mdAllBtn']) assert.ok(ids.includes(`'${id}'`), id);
  assert.match(js, /setDeckControls\(false\);\s*vscode\.postMessage\(\{ type: 'ready' \}\);/);
  // Slides are drawn one at a time: load() once, then renderSlide(i) on first use.
  assert.match(js, /previewer\.load\(buf\)/);
  assert.doesNotMatch(js, /previewer\.preview\(/);
});

test('pollLocalFile reports a change made by another program, and stops after dispose', async () => {
  const Module = require('module');
  const orig = Module._load;
  Module._load = function (req, ...rest) { return req === 'vscode' ? {} : orig.call(this, req, ...rest); };
  let ext;
  try { ext = require('../src/extension'); } finally { Module._load = orig; }
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx4-'));
  const file = path.join(dir, 'deck.pptx');
  fs.writeFileSync(file, 'a');
  let n = 0;
  const d = ext.pollLocalFile({ scheme: 'file', fsPath: file }, () => { n++; }, 50);
  try {
    await new Promise((r) => setTimeout(r, 120));
    fs.writeFileSync(file, 'changed by a script');
    await new Promise((r) => setTimeout(r, 300));
    assert.ok(n >= 1, 'change seen');
    d.dispose();
    const after = n;
    fs.writeFileSync(file, 'changed again after dispose');
    await new Promise((r) => setTimeout(r, 300));
    assert.equal(n, after);
    assert.doesNotThrow(() => ext.pollLocalFile({ scheme: 'vscode-remote', path: '/x' }, () => {}).dispose());
  } finally {
    d.dispose();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('protocol: savePng takes only a real PNG of sane size for a real slide number', () => {
  const { parseWebviewMessage, isPng } = require('../src/protocol');
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]);
  assert.ok(isPng(png));
  const ok = parseWebviewMessage({ type: 'savePng', slide: 2, data: png });
  assert.equal(ok.slide, 2);
  assert.equal(ok.data, png);
  assert.equal(parseWebviewMessage({ type: 'savePng', slide: 2, data: png.buffer }).data.length, png.length);
  assert.equal(parseWebviewMessage({ type: 'savePng', slide: 2, data: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]) }), null);
  assert.equal(parseWebviewMessage({ type: 'savePng', slide: 0, data: png }), null);
  assert.equal(parseWebviewMessage({ type: 'savePng', slide: 2, data: 'iVBORw0KGgo=' }), null);
  assert.equal(parseWebviewMessage({ type: 'savePng', slide: 2, data: Array.from(png) }), null);
  assert.deepEqual(parseWebviewMessage({ type: 'pngDone', slide: 3 }), { type: 'pngDone', slide: 3 });
  assert.equal(parseWebviewMessage({ type: 'pngFailed', message: 'x'.repeat(900) }).message.length, 300);
});

test('package.json: PNG commands, phrase-first name with every word of the search phrases, 0.4.3', () => {
  const pkg = require('../package.json');
  const ids = pkg.contributes.commands.map((c) => c.command);
  assert.ok(ids.includes('pptxViewer.copySlidePng') && ids.includes('pptxViewer.saveSlidePng'));
  assert.ok(pkg.displayName.startsWith('PPTX Viewer'));
  const words = new Set(pkg.displayName.toLowerCase().split(/[^a-z0-9]+/));
  for (const phrase of ['pptx viewer', 'ppt viewer', 'powerpoint viewer', 'powerpoint preview', 'ppt preview', 'presentation viewer']) assert.ok(phrase.split(' ').every((w) => words.has(w)), phrase);
  assert.ok(pkg.keywords.length <= 30 && new Set(pkg.keywords).size === pkg.keywords.length);
  assert.equal(pkg.version, '0.4.3');
  const { WEBVIEW_COMMANDS } = (() => { const M = require('module'); const o = M._load; M._load = function (r, ...a) { return r === 'vscode' ? {} : o.call(this, r, ...a); }; try { return require('../src/extension'); } finally { M._load = o; } })();
  assert.ok(WEBVIEW_COMMANDS.includes('copySlidePng') && WEBVIEW_COMMANDS.includes('saveSlidePng'));
});

test('a shape path with NaN coordinates is written as an empty path (no console error per shape)', () => {
  const js = fs.readFileSync(path.join(ROOT, 'media/viewer.js'), 'utf8');
  const m = js.match(/Element\.prototype\.setAttribute = function \(name, value\) \{([^}]+)\}/);
  assert.ok(m);
  const calls = [];
  const fake = { setAttribute(n, v) { calls.push([n, v]); } };
  const fn = new Function('setAttr', `return function (name, value) {${m[1]}}`)((n, v) => calls.push([n, v]));
  fn.call(fake, 'd', 'M NaN,NaN L 1,2');
  fn.call(fake, 'd', 'M 0,0 L 1,2');
  fn.call(fake, 'title', 'NaN');
  assert.deepEqual(calls, [['d', 'M0,0'], ['d', 'M 0,0 L 1,2'], ['title', 'NaN']]);
});
