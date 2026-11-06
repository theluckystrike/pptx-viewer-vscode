'use strict';
// Webview UI. All text from the deck is inserted with textContent or text nodes, never as HTML.
// Slide drawing is done by pptx-preview into a hidden host; we clone its slide nodes.
(function () {
  const vscode = acquireVsCodeApi();
  const C = window.PptxCore;
  const $ = (id) => document.getElementById(id);
  const el = (tag, cls, text) => {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  };

  const state = {
    deck: null, fileName: '', current: 1, query: '', hits: [], hitPos: -1,
    drawn: null, drawError: '', config: { showNotes: true, renderSlides: true }, notes: true,
  };
  const saved = vscode.getState();
  if (saved && Number.isInteger(saved.current)) state.current = saved.current;

  function status(text) { $('status').textContent = text || ''; $('status').hidden = !text; }

  // ---- messages from the extension host ----
  window.addEventListener('message', (ev) => {
    const m = ev.data;
    if (!m || typeof m !== 'object') return;
    if (m.type === 'init') {
      state.config = m.config || state.config;
      state.notes = state.config.showNotes;
      state.fileName = String(m.fileName || '');
      $('fileName').textContent = state.fileName;
      status('Reading ' + state.fileName + '…');
      applyNotes();
    } else if (m.type === 'config') {
      state.config = m.config || state.config;
    } else if (m.type === 'deck') {
      onDeck(m.deck);
    } else if (m.type === 'bytes') {
      draw(m.data);
    } else if (m.type === 'tooLarge') {
      showTooLarge(m.size, m.limit);
    } else if (m.type === 'error') {
      status('');
      showError(String(m.message || 'Unknown error'));
    } else if (m.type === 'command') {
      if (m.name === 'find') { $('searchInput').focus(); $('searchInput').select(); }
      else if (m.name === 'toggleNotes') { state.notes = !state.notes; applyNotes(); }
    }
  });

  function showError(text) {
    const box = $('notice');
    box.replaceChildren(el('strong', null, 'Cannot open this file. '), document.createTextNode(text));
    box.hidden = false;
    box.className = 'error';
  }

  function showTooLarge(size, limit) {
    status('');
    const box = $('notice');
    const btn = el('button', null, 'Open anyway');
    btn.addEventListener('click', () => { box.hidden = true; status('Reading…'); vscode.postMessage({ type: 'loadAnyway' }); });
    box.replaceChildren(document.createTextNode(`This deck is ${C.formatBytes(size)}, above the ${C.formatBytes(limit)} limit (setting pptxViewer.maxFileSizeMB). `), btn);
    box.className = 'warn';
    box.hidden = false;
  }

  function onDeck(deck) {
    state.deck = deck;
    state.drawn = null;
    state.drawError = '';
    const n = deck.slides.length;
    state.current = C.clampSlide(state.current, n);
    $('slideCount').textContent = '/ ' + n;
    if (!n) { status(''); showError('The deck has no slides.'); return; }
    buildThumbs();
    buildOutline();
    status(state.config.renderSlides ? `Drawing ${n} slides…` : '');
    if (deck.hasMacros) note('This deck contains macros. They are never run.');
    go(state.current, true);
  }

  function note(text) {
    const s = $('status');
    s.textContent = text;
    s.hidden = false;
  }

  // ---- drawing with pptx-preview ----
  async function draw(data) {
    const deck = state.deck;
    if (!deck || typeof pptxPreview === 'undefined') { state.drawError = 'The slide renderer did not load.'; renderStage(); return; }
    const t0 = performance.now();
    const host = $('renderHost');
    host.replaceChildren();
    const box = el('div');
    host.appendChild(box);
    try {
      const buf = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
      const previewer = pptxPreview.init(box, { width: deck.widthPx, height: deck.heightPx, mode: 'list' });
      await previewer.preview(buf);
      const nodes = Array.from(box.querySelectorAll('.pptx-preview-slide-wrapper'));
      state.drawn = nodes;
      status(deck.hasMacros ? 'This deck contains macros. They are never run.' : '');
      vscode.postMessage({ type: 'rendered', slides: nodes.length, ms: Math.round(performance.now() - t0) });
    } catch (e) {
      state.drawn = null;
      state.drawError = 'Could not draw the slides (' + (e && e.message ? e.message : String(e)) + '). Text, notes and search still work.';
      status('');
      vscode.postMessage({ type: 'error', message: state.drawError });
    }
    refreshThumbs();
    renderStage();
  }

  function slideClone(i, targetW, targetH) {
    const src = state.drawn && state.drawn[i - 1];
    if (!src) return null;
    const deck = state.deck;
    const scale = C.fitScale(targetW, targetH, deck.widthPx, deck.heightPx);
    const frame = el('div', 'slideFrame');
    frame.style.width = Math.floor(deck.widthPx * scale) + 'px';
    frame.style.height = Math.floor(deck.heightPx * scale) + 'px';
    const inner = src.cloneNode(true);
    inner.classList.add('slideInner');
    inner.style.transform = `scale(${scale})`;
    inner.style.width = deck.widthPx + 'px';
    inner.style.height = deck.heightPx + 'px';
    frame.appendChild(inner);
    return frame;
  }

  // Text view of a slide, used before drawing finishes, when drawing is off, or when it fails.
  function textSlide(slide) {
    const card = el('div', 'textSlide');
    card.appendChild(hl('h2', slide.title || slide.fallbackTitle || 'Slide ' + slide.index));
    for (const b of slide.blocks) {
      if (b.kind === 'text') {
        const ul = el('ul');
        for (const p of b.paragraphs) { const li = hl('li', p.text); li.style.marginLeft = p.level * 1.4 + 'em'; ul.appendChild(li); }
        card.appendChild(ul);
      } else if (b.kind === 'table') {
        const t = el('table');
        b.rows.forEach((r, ri) => { const tr = el('tr'); r.forEach((c) => tr.appendChild(hl(ri ? 'td' : 'th', c))); t.appendChild(tr); });
        card.appendChild(t);
      }
    }
    if (slide.images) card.appendChild(el('p', 'muted', `${slide.images} picture${slide.images > 1 ? 's' : ''} on this slide.`));
    return card;
  }

  // Element whose text has the search matches wrapped in <mark>.
  function hl(tag, text, cls) {
    const e = el(tag, cls);
    for (const seg of C.splitHighlight(text, state.query)) {
      if (seg.hit) e.appendChild(el('mark', null, seg.text));
      else e.appendChild(document.createTextNode(seg.text));
    }
    return e;
  }

  function renderStage() {
    const deck = state.deck;
    if (!deck) return;
    const slide = deck.slides[state.current - 1];
    const stage = $('stage');
    const wrap = $('stageWrap');
    const notesH = state.notes ? $('notes').offsetHeight : 0;
    const w = wrap.clientWidth - 32;
    const h = Math.max(160, wrap.clientHeight - notesH - 48 - ($('notice').hidden ? 0 : $('notice').offsetHeight));
    const clone = slideClone(state.current, w, h);
    stage.replaceChildren(clone || textSlide(slide));
    if (slide.hidden) stage.appendChild(el('div', 'badge', 'Hidden slide'));
    const msg = state.drawError || C.noticeMessage(slide.notices);
    const box = $('notice');
    if (msg) { box.textContent = msg; box.className = state.drawError ? 'error' : 'info'; box.hidden = false; }
    else if (box.className !== 'warn') box.hidden = true;
    const notes = $('notes');
    notes.replaceChildren(el('div', 'notesLabel', 'Speaker notes'));
    if (slide.notes) for (const line of slide.notes.split('\n')) notes.appendChild(hl('p', line));
    else notes.appendChild(el('p', 'muted', 'No notes on this slide.'));
  }

  // ---- thumbnails (built lazily as they scroll into view) ----
  let observer = null;
  function buildThumbs() {
    const list = $('thumbs');
    list.replaceChildren();
    if (observer) observer.disconnect();
    observer = new IntersectionObserver((entries) => {
      for (const e of entries) if (e.isIntersecting) fillThumb(e.target);
    }, { root: list, rootMargin: '400px' });
    for (const s of state.deck.slides) {
      const item = el('div', 'thumb' + (s.hidden ? ' hiddenSlide' : ''));
      item.setAttribute('role', 'listitem');
      item.tabIndex = 0;
      item.dataset.index = String(s.index);
      item.appendChild(el('span', 'num', String(s.index)));
      const pic = el('div', 'pic');
      pic.style.aspectRatio = `${state.deck.widthPx} / ${state.deck.heightPx}`;
      pic.appendChild(el('span', 'thumbTitle', s.title || s.fallbackTitle || ''));
      item.appendChild(pic);
      item.addEventListener('click', () => go(s.index));
      list.appendChild(item);
      observer.observe(item);
    }
  }

  function fillThumb(item) {
    if (!state.drawn || item.dataset.filled) return;
    const pic = item.querySelector('.pic');
    const clone = slideClone(parseInt(item.dataset.index, 10), pic.clientWidth || 160, 1000);
    if (!clone) return;
    pic.replaceChildren(clone);
    item.dataset.filled = '1';
  }

  function refreshThumbs() {
    if (!observer) return;
    for (const item of $('thumbs').children) { observer.unobserve(item); observer.observe(item); }
  }

  // ---- outline and search ----
  function buildOutline() {
    const box = $('outline');
    box.replaceChildren();
    const deck = state.deck;
    if (state.query) {
      const hits = state.hits;
      box.appendChild(el('div', 'muted', hits.length ? `${hits.length} slide${hits.length > 1 ? 's' : ''} match` : 'No slide matches.'));
      for (const h of hits) {
        const s = deck.slides[h.index - 1];
        const a = el('a', 'outlineItem' + (h.index === state.current ? ' current' : ''));
        a.href = '#';
        a.appendChild(el('span', 'num', String(h.index)));
        a.appendChild(hl('span', s.title || s.fallbackTitle || 'Slide ' + s.index, 'oTitle'));
        a.appendChild(hl('span', (h.where === 'notes' ? 'Notes: ' : '') + h.snippet, 'oSnippet'));
        a.addEventListener('click', (e) => { e.preventDefault(); go(h.index); });
        box.appendChild(a);
      }
      return;
    }
    for (const s of deck.slides) {
      const a = el('a', 'outlineItem' + (s.index === state.current ? ' current' : ''));
      a.href = '#';
      a.appendChild(el('span', 'num', String(s.index)));
      a.appendChild(el('span', 'oTitle', s.title || s.fallbackTitle || 'Slide ' + s.index));
      const lines = C.slideBodyLines(s).slice(0, 4);
      for (const l of lines) a.appendChild(el('span', 'oSnippet', l));
      a.addEventListener('click', (e) => { e.preventDefault(); go(s.index); });
      box.appendChild(a);
    }
  }

  function setTab(name) {
    $('tabSlides').classList.toggle('active', name === 'slides');
    $('tabOutline').classList.toggle('active', name === 'outline');
    $('thumbs').hidden = name !== 'slides';
    $('outline').hidden = name !== 'outline';
  }

  function runSearch() {
    state.query = $('searchInput').value.trim();
    state.hits = C.searchDeck(state.deck, state.query);
    state.hitPos = -1;
    $('searchCount').textContent = state.query ? (state.hits.length ? `${state.hits.length} slide${state.hits.length > 1 ? 's' : ''}` : 'No match') : '';
    if (state.query) setTab('outline');
    buildOutline();
    renderStage();
  }

  function nextHit(dir) {
    if (!state.hits.length) return;
    state.hitPos = (state.hitPos + dir + state.hits.length) % state.hits.length;
    $('searchCount').textContent = `${state.hitPos + 1} of ${state.hits.length} slide${state.hits.length > 1 ? 's' : ''}`;
    go(state.hits[state.hitPos].index);
  }

  // ---- navigation ----
  function go(n, force) {
    if (!state.deck) return;
    n = C.clampSlide(n, state.deck.slides.length);
    const changed = n !== state.current;
    state.current = n;
    $('slideInput').value = String(n);
    for (const t of $('thumbs').children) {
      const on = t.dataset.index === String(n);
      t.classList.toggle('current', on);
      if (on && (changed || force)) t.scrollIntoView({ block: 'nearest' });
    }
    for (const a of $('outline').querySelectorAll('.outlineItem')) a.classList.toggle('current', a.querySelector('.num').textContent === String(n));
    renderStage();
    vscode.setState({ current: n });
    if (changed || force) vscode.postMessage({ type: 'slideChanged', slide: n });
  }

  function applyNotes() {
    $('notes').hidden = !state.notes;
    $('notesBtn').classList.toggle('on', state.notes);
    renderStage();
  }

  $('prevBtn').addEventListener('click', () => go(state.current - 1));
  $('nextBtn').addEventListener('click', () => go(state.current + 1));
  $('slideInput').addEventListener('change', () => go($('slideInput').value));
  $('tabSlides').addEventListener('click', () => setTab('slides'));
  $('tabOutline').addEventListener('click', () => setTab('outline'));
  $('sidebarBtn').addEventListener('click', () => document.body.classList.toggle('noSidebar'));
  $('notesBtn').addEventListener('click', () => { state.notes = !state.notes; applyNotes(); });
  $('mdBtn').addEventListener('click', () => vscode.postMessage({ type: 'copySlideMarkdown', slide: state.current }));
  $('mdAllBtn').addEventListener('click', () => vscode.postMessage({ type: 'copyDeckMarkdown' }));
  let searchTimer = null;
  $('searchInput').addEventListener('input', () => { clearTimeout(searchTimer); searchTimer = setTimeout(runSearch, 120); });
  $('searchInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); clearTimeout(searchTimer); if (state.query !== $('searchInput').value.trim()) runSearch(); nextHit(e.shiftKey ? -1 : 1); }
    else if (e.key === 'Escape') { $('searchInput').value = ''; runSearch(); setTab('slides'); }
  });
  document.addEventListener('keydown', (e) => {
    if (e.target instanceof HTMLInputElement) return;
    if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') { e.preventDefault(); go(state.current + 1); }
    else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { e.preventDefault(); go(state.current - 1); }
    else if (e.key === 'Home') go(1);
    else if (e.key === 'End' && state.deck) go(state.deck.slides.length);
  });
  // Links inside drawn slides must not navigate the webview.
  document.addEventListener('click', (e) => { const a = e.target.closest && e.target.closest('#stage a, #thumbs a'); if (a) e.preventDefault(); }, true);
  let resizeTimer = null;
  window.addEventListener('resize', () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(renderStage, 100); });

  vscode.postMessage({ type: 'ready' });
})();
