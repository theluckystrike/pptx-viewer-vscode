'use strict';
// Old binary PowerPoint 97-2003 files read as text. Fixtures: two sample decks from the Apache POI
// test data (Apache-2.0, see THIRD-PARTY-NOTICES.md), a .pptx renamed to .ppt and a cut-off file.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { parsePpt, sniffPpt, openCfb } = require('../src/ppt');
const { slideToMarkdown, deckToMarkdown, slideSearchText } = require('../src/pptx');

const read = (f) => fs.readFileSync(path.join(__dirname, 'fixtures', f));

test('slide titles, body text and speaker notes in slide order', () => {
  const d = parsePpt(read('legacy-basic.ppt'));
  assert.equal(d.kind, 'ppt');
  assert.equal(d.legacy, true);
  assert.equal(d.slides.length, 2);
  assert.equal(d.slides[0].title, 'This is a test title');
  assert.deepEqual(d.slides[0].blocks[0].paragraphs.map((p) => p.text), ['This is a test subtitle', 'This is on page 1']);
  assert.equal(d.slides[0].notes, 'These are the notes for page 1');
  assert.equal(d.slides[1].title, 'This is the title on page 2');
  assert.deepEqual(d.slides[1].blocks[0].paragraphs.map((p) => p.text), ['This is page two', 'It has several blocks of text', 'None of them have formatting']);
  assert.equal(d.slides[1].notes, 'These are the notes on page two, again lacking formatting');
});

test('free text boxes (not placeholders) are read from the slide itself', () => {
  const d = parsePpt(read('legacy-textbox.ppt'));
  assert.equal(d.slides.length, 1);
  const texts = d.slides[0].blocks.flatMap((b) => b.paragraphs.map((p) => p.text));
  assert.deepEqual(texts, ['Hello, World!!!', 'I am just a poor boy', 'This is Times New Roman', 'Plain Text']);
  assert.equal(d.slides[0].fallbackTitle, 'Hello, World!!!');
});

test('search text and Copy as Markdown work on the old format', () => {
  const d = parsePpt(read('legacy-basic.ppt'));
  assert.match(slideSearchText(d.slides[1]), /several blocks/);
  const md = slideToMarkdown(d.slides[0]);
  assert.match(md, /^## This is a test title/);
  assert.match(md, /> \*\*Notes:\*\* These are the notes for page 1/);
  assert.match(deckToMarkdown(d, 'legacy-basic.ppt'), /This is the title on page 2/);
});

test('the first bytes decide the format, not the name', () => {
  assert.equal(sniffPpt(read('legacy-basic.ppt')), 'ole');
  assert.equal(sniffPpt(read('renamed-pptx.ppt')), 'zip');
  assert.equal(sniffPpt(Buffer.from('{\\rtf1 hello}')), 'other');
});

test('a cut-off or foreign OLE file throws a readable message, never an empty deck', () => {
  assert.throws(() => parsePpt(read('legacy-broken.ppt')), /cannot be read|damaged/);
  // An OLE file without a PowerPoint stream: take the .ppt and rename its stream entry.
  const b = Buffer.from(read('legacy-basic.ppt'));
  const name = Buffer.from('PowerPoint Document', 'utf16le');
  const at = b.indexOf(name);
  assert.ok(at > 0);
  b.write('X', at, 'utf16le');
  assert.throws(() => parsePpt(b), /not a PowerPoint 97-2003 deck/);
  assert.throws(() => openCfb(Buffer.alloc(1024)), /not an OLE file/);
});
