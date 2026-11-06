'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { zipSync, strToU8 } = require('fflate');
const { parseDeck, slideToMarkdown, deckToMarkdown, slideSearchText } = require('../src/pptx');
const { parseXml, textOf, find } = require('../src/ooxml/xml');
const { resolveTarget, relsPathFor, readXmlParts } = require('../src/ooxml/zip');

const fx = (f) => fs.readFileSync(path.join(__dirname, 'fixtures', f));

test('quarterly-review: order, titles, levels, table, notes, notices, size', () => {
  const d = parseDeck(fx('quarterly-review.pptx'));
  assert.equal(d.kind, 'pptx');
  assert.equal(d.slides.length, 6);
  assert.deepEqual([d.widthPx, d.heightPx], [1280, 720]);
  assert.deepEqual(d.slides.map((s) => s.title), ['Q3 Platform Review', 'Summary', 'Latency by region', 'Where the time goes', 'Requests per day', 'Next steps']);
  const bullets = d.slides[1].blocks[0].paragraphs;
  assert.equal(bullets[1].text, 'Cache hit rate is now 93%');
  assert.equal(bullets[1].level, 1);
  const table = d.slides[2].blocks.find((b) => b.kind === 'table');
  assert.deepEqual(table.rows[0], ['Region', 'p50 (ms)', 'p95 (ms)', 'Change']);
  assert.equal(table.rows.length, 4);
  assert.match(d.slides[0].notes, /Keep the intro under two minutes/);
  assert.deepEqual(d.slides[4].notices, ['chart']);
  assert.equal(d.slides[3].images, 1);
  assert.equal(d.hasMacros, false);
});

test('ppsx and potx are detected from the content type', () => {
  assert.equal(parseDeck(fx('quarterly-review.ppsx')).kind, 'ppsx');
  const t = parseDeck(fx('template.potx'));
  assert.equal(t.kind, 'potx');
  assert.equal(t.slides.length, 6);
});

test('SmartArt and video are reported, not drawn', () => {
  const d = parseDeck(fx('smartart.pptx'));
  assert.deepEqual(d.slides[0].notices.sort(), ['smartart', 'video']);
});

test('120 slides parse fast and keep order', () => {
  const t = Date.now();
  const d = parseDeck(fx('many-slides.pptx'));
  assert.equal(d.slides.length, 120);
  assert.equal(d.slides[119].title, 'Slide 120');
  assert.equal(d.slides[9].notes, 'Checkpoint 10');
  assert.ok(Date.now() - t < 2000);
});

test('hostile text stays text and is escaped in Markdown', () => {
  const d = parseDeck(fx('hostile.pptx'));
  assert.equal(d.slides[0].title, 'Title <img src=x onerror="document.title=\'PWNED\'">');
  const md = slideToMarkdown(d.slides[0]);
  assert.doesNotMatch(md, /(^|[^\\])</);
  assert.ok(md.includes('\\<img'));
  assert.ok(md.includes('\\<script\\>'));
});

test('Markdown: heading, nested bullets, table, notes', () => {
  const d = parseDeck(fx('quarterly-review.pptx'));
  const md = slideToMarkdown(d.slides[1]);
  assert.ok(md.startsWith('## Summary\n'));
  assert.ok(md.includes('\n  - Cache hit rate is now 93%'));
  assert.ok(md.includes('> **Notes:** Stress the latency win first.'));
  const tmd = slideToMarkdown(d.slides[2]);
  assert.ok(tmd.includes('| Region | p50 (ms) | p95 (ms) | Change |\n| --- | --- | --- | --- |'));
  assert.ok(slideToMarkdown(d.slides[4]).includes('Not shown as text: chart'));
  const all = deckToMarkdown(d, 'q.pptx');
  assert.ok(all.startsWith('# q.pptx\n'));
  assert.equal(all.match(/^## /gm).length, 6);
});

test('search text includes notes and table cells', () => {
  const d = parseDeck(fx('quarterly-review.pptx'));
  assert.match(slideSearchText(d.slides[2]), /ap-south 95 240/);
  assert.match(slideSearchText(d.slides[5]), /owners/);
});

test('not a zip: clear error (old .ppt or damaged)', () => {
  assert.throws(() => parseDeck(Buffer.from('\xd0\xcf\x11\xe0 old binary ppt', 'latin1')), /not a ZIP package/);
});

test('zip without a presentation part: clear error', () => {
  const z = zipSync({ 'hello.xml': strToU8('<a/>') });
  assert.throws(() => parseDeck(z), /No presentation part/);
});

test('media parts are not unzipped', () => {
  const z = zipSync({ 'a.xml': strToU8('<a/>'), 'ppt/media/big.png': new Uint8Array(1000) });
  assert.deepEqual([...readXmlParts(z).keys()], ['a.xml']);
});

test('hidden slide and slide without title', () => {
  const ct = '<Types><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>';
  const pres = '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>';
  const rels = '<Relationships><Relationship Id="rId1" Type="http://x/slide" Target="slides/slide1.xml"/></Relationships>';
  const slide = '<p:sld show="0"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Only body &amp; text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>';
  const z = zipSync({ '[Content_Types].xml': strToU8(ct), 'ppt/presentation.xml': strToU8(pres), 'ppt/_rels/presentation.xml.rels': strToU8(rels), 'ppt/slides/slide1.xml': strToU8(slide) });
  const d = parseDeck(z);
  assert.equal(d.slides[0].hidden, true);
  assert.equal(d.slides[0].title, '');
  assert.equal(d.slides[0].fallbackTitle, 'Only body & text');
  assert.deepEqual([d.widthPx, d.heightPx], [1280, 720]);
});

test('xml: entities, CDATA, comments, self-closing, quotes with >', () => {
  const doc = parseXml('<?xml version="1.0"?><!-- c --><r a="x &gt; y" b=\'1\'><t>A &amp; B &#65;&#x42;</t><![CDATA[<raw>]]><e/></r>');
  const r = find(doc, 'r');
  assert.equal(r.attrs.a, 'x > y');
  assert.equal(r.attrs.b, '1');
  assert.equal(textOf(find(doc, 't')), 'A & B AB');
  assert.ok(textOf(r).includes('<raw>'));
  assert.equal(parseXml('<a>&#xFFFFFFF;</a>').children[0].children[0], '');
});

test('zip paths: resolve relationship targets', () => {
  assert.equal(resolveTarget('ppt/presentation.xml', 'slides/slide1.xml'), 'ppt/slides/slide1.xml');
  assert.equal(resolveTarget('ppt/slides/slide1.xml', '../notesSlides/notesSlide1.xml'), 'ppt/notesSlides/notesSlide1.xml');
  assert.equal(resolveTarget('ppt/slides/slide1.xml', '/ppt/media/a.png'), 'ppt/media/a.png');
  assert.equal(relsPathFor('ppt/slides/slide1.xml'), 'ppt/slides/_rels/slide1.xml.rels');
});
