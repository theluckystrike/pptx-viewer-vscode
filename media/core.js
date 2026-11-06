'use strict';
// Pure helpers for the webview. No DOM access, so node --test can load this file.
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PptxCore = api;
})(typeof self !== 'undefined' ? self : this, function () {
  function norm(s) { return String(s || '').toLocaleLowerCase(); }

  function slideBodyLines(slide) {
    const out = [];
    for (const b of slide.blocks || []) {
      if (b.kind === 'text') for (const p of b.paragraphs) out.push(p.text);
      else if (b.kind === 'table') for (const r of b.rows) out.push(r.join(' | '));
    }
    return out;
  }

  /** Split text into [{text, hit}] segments for every case-insensitive match of query. */
  function splitHighlight(text, query) {
    text = String(text || '');
    const q = norm(query);
    if (!q) return [{ text, hit: false }];
    const low = norm(text);
    const out = [];
    let i = 0;
    for (;;) {
      const j = low.indexOf(q, i);
      if (j < 0 || low.length !== text.length) {
        // Lower-casing changed the length (rare scripts): fall back to no highlight.
        if (j >= 0 && low.length !== text.length) return [{ text, hit: false }];
        break;
      }
      if (j > i) out.push({ text: text.slice(i, j), hit: false });
      out.push({ text: text.slice(j, j + q.length), hit: true });
      i = j + q.length;
    }
    if (i < text.length || !out.length) out.push({ text: text.slice(i), hit: false });
    return out;
  }

  function countMatches(text, query) {
    const q = norm(query);
    if (!q) return 0;
    const low = norm(text);
    let n = 0;
    for (let i = low.indexOf(q); i >= 0; i = low.indexOf(q, i + q.length)) n++;
    return n;
  }

  function snippet(line, query, width) {
    width = width || 90;
    const j = norm(line).indexOf(norm(query));
    if (j < 0 || line.length <= width) return line;
    const start = Math.max(0, j - Math.floor(width / 3));
    return (start > 0 ? '…' : '') + line.slice(start, start + width) + (start + width < line.length ? '…' : '');
  }

  /** Search titles, body text, tables and notes. Returns one hit per matching slide. */
  function searchDeck(deck, query) {
    const q = String(query || '').trim();
    if (!q || !deck) return [];
    const hits = [];
    for (const s of deck.slides) {
      const title = s.title || s.fallbackTitle || '';
      const body = slideBodyLines(s);
      const notes = s.notes ? s.notes.split('\n') : [];
      const count = countMatches(title, q) + body.reduce((a, l) => a + countMatches(l, q), 0) + notes.reduce((a, l) => a + countMatches(l, q), 0);
      if (!count) continue;
      let where = 'title';
      let line = title;
      if (!countMatches(title, q)) {
        const b = body.find((l) => countMatches(l, q));
        if (b !== undefined) { where = 'text'; line = b; } else { where = 'notes'; line = notes.find((l) => countMatches(l, q)) || ''; }
      }
      hits.push({ index: s.index, count, where, snippet: snippet(line, q) });
    }
    return hits;
  }

  const NOTICE_NAMES = {
    smartart: 'SmartArt', chart: 'a chart', video: 'a video', audio: 'audio', object: 'an embedded object', ink: 'an ink drawing', missing: 'a missing part',
  };

  function joinList(items) {
    if (items.length <= 1) return items.join('');
    return items.slice(0, -1).join(', ') + ' and ' + items[items.length - 1];
  }

  /** Clear message for content we cannot draw, or '' when there is none. */
  function noticeMessage(notices) {
    if (!notices || !notices.length) return '';
    const names = notices.map((n) => NOTICE_NAMES[n] || n);
    const what = joinList(names);
    const parts = ['This slide has ' + what + '.'];
    if (notices.includes('chart')) parts.push('Charts are drawn from the numbers saved in the file and can look different from PowerPoint.');
    const cannot = notices.filter((n) => n !== 'chart').map((n) => NOTICE_NAMES[n] || n);
    if (cannot.length) parts.push('PPTX Slide Viewer cannot draw ' + joinList(cannot) + '. Open the deck in PowerPoint, Keynote or LibreOffice to see ' + (cannot.length > 1 ? 'them' : 'it') + '.');
    return parts.join(' ');
  }

  function fitScale(boxW, boxH, w, h) {
    if (!(boxW > 0 && boxH > 0 && w > 0 && h > 0)) return 1;
    return Math.min(boxW / w, boxH / h);
  }

  function clampSlide(n, count) {
    n = parseInt(n, 10);
    if (!Number.isFinite(n)) return 1;
    return Math.max(1, Math.min(count || 1, n));
  }

  function formatBytes(n) {
    if (n >= 1024 * 1024) return (n / 1024 / 1024).toFixed(1) + ' MB';
    if (n >= 1024) return Math.round(n / 1024) + ' KB';
    return n + ' B';
  }

  return { splitHighlight, countMatches, searchDeck, noticeMessage, fitScale, clampSlide, formatBytes, slideBodyLines, snippet };
});
