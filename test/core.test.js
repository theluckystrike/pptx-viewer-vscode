'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const C = require('../media/core');
const { parseDeck } = require('../src/pptx');

const deck = parseDeck(fs.readFileSync(path.join(__dirname, 'fixtures', 'quarterly-review.pptx')));

test('searchDeck finds titles, body, tables and notes, case-insensitive', () => {
  assert.deepEqual(C.searchDeck(deck, 'LATENCY').map((h) => h.index), [1, 2, 3, 6]);
  const notes = C.searchDeck(deck, 'owners');
  assert.equal(notes.length, 1);
  assert.equal(notes[0].where, 'notes');
  assert.equal(C.searchDeck(deck, 'ap-south')[0].where, 'text');
  assert.deepEqual(C.searchDeck(deck, '   '), []);
  assert.deepEqual(C.searchDeck(deck, 'zzzz'), []);
});

test('splitHighlight keeps all text and marks every match', () => {
  const segs = C.splitHighlight('Latency and latency', 'latency');
  assert.equal(segs.map((s) => s.text).join(''), 'Latency and latency');
  assert.equal(segs.filter((s) => s.hit).length, 2);
  assert.deepEqual(C.splitHighlight('abc', ''), [{ text: 'abc', hit: false }]);
  assert.deepEqual(C.splitHighlight('<b>x</b>', '<b>'), [{ text: '<b>', hit: true }, { text: 'x</b>', hit: false }]);
});

test('noticeMessage names what cannot be drawn', () => {
  assert.equal(C.noticeMessage([]), '');
  assert.match(C.noticeMessage(['smartart', 'video']), /cannot draw SmartArt and a video/);
  assert.match(C.noticeMessage(['chart']), /drawn from the numbers saved in the file/);
  assert.doesNotMatch(C.noticeMessage(['chart']), /cannot draw/);
});

test('fitScale, clampSlide, formatBytes', () => {
  assert.equal(C.fitScale(640, 1000, 1280, 720), 0.5);
  assert.equal(C.fitScale(0, 100, 10, 10), 1);
  assert.equal(C.clampSlide('9', 6), 6);
  assert.equal(C.clampSlide('x', 6), 1);
  assert.equal(C.clampSlide(0, 6), 1);
  assert.equal(C.formatBytes(2 * 1024 * 1024), '2.0 MB');
});
