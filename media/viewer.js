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
    view: 'slide', zoom: 'fit', scale: 1,
  };
  const saved = vscode.getState();
  if (saved && Number.isInteger(saved.current)) state.current = saved.current;

  function saveState() {
    vscode.setState({ current: state.current, view: state.view, zoom: state.zoom });
  }

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
      // Order: this panel's own state (a reload), then the saved choice for this file, then the settings.
      const prefs = (saved && saved.view) ? saved : (m.prefs || {});
      state.view = C.normalizeView(prefs.view || state.config.defaultView);
      state.zoom = C.normalizeZoom(prefs.zoom !== undefined ? prefs.zoom : state.config.defaultZoom);
      applyViewButtons();
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
      else if (m.name === 'zoomIn') zoomStep(1);
      else if (m.name === 'zoomOut') zoomStep(-1);
      else if (m.name === 'zoomFit') setZoom('fit');
      else if (m.name === 'zoomWidth') setZoom('width');
      else if (m.name === 'zoomActual') setZoom(1);
      else if (m.name === 'toggleScrollView') setView(state.view === 'scroll' ? 'slide' : 'scroll');
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

  function slideClone(i, scale) {
    const src = state.drawn && state.drawn[i - 1];
    if (!src) return null;
    const deck = state.deck;
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

  // Size of the box a slide is fitted into, inside the stage.
  function stageBox() {
    const wrap = $('stageWrap');
    const notesH = state.notes ? $('notes').offsetHeight : 0;
    const w = Math.max(80, wrap.clientWidth - 32);
    const h = Math.max(160, wrap.clientHeight - notesH - 48 - ($('notice').hidden ? 0 : $('notice').offsetHeight));
    return { w, h };
  }

  function currentScale() {
    const deck = state.deck;
    if (!deck) return 1;
    const b = stageBox();
    // In the scroll view a vertical scroll bar takes some width; keep Fit and Fit width inside it,
    // so the column never gets a horizontal scroll bar of a few pixels.
    const w = state.view === 'scroll' ? b.w - 24 : b.w;
    return C.zoomScale(state.zoom, w, b.h, deck.widthPx, deck.heightPx);
  }

  function renderStage() {
    const deck = state.deck;
    if (!deck) return;
    const slide = deck.slides[state.current - 1];
    const stage = $('stage');
    state.scale = currentScale();
    updateZoomControls();
    if (state.view === 'scroll') renderScrollList();
    else {
      const clone = slideClone(state.current, state.scale);
      stage.classList.remove('scrollView');
      stage.replaceChildren(clone || textSlide(slide));
      if (slide.hidden) stage.appendChild(el('div', 'badge', 'Hidden slide'));
    }
    renderSlideInfo();
  }

  // Notice and notes for the current slide. Used on its own while the scroll view scrolls.
  function renderSlideInfo() {
    const slide = state.deck.slides[state.current - 1];
    const msg = state.drawError || C.noticeMessage(slide.notices);
    const box = $('notice');
    if (msg) { box.textContent = msg; box.className = state.drawError ? 'error' : 'info'; box.hidden = false; }
    else if (box.className !== 'warn') box.hidden = true;
    const notes = $('notes');
    const label = state.view === 'scroll' ? 'Speaker notes, slide ' + slide.index : 'Speaker notes';
    notes.replaceChildren(el('div', 'notesLabel', label));
    if (slide.notes) for (const line of slide.notes.split('\n')) notes.appendChild(hl('p', line));
    else notes.appendChild(el('p', 'muted', 'No notes on this slide.'));
  }

  // ---- scroll view: every slide in one column, drawn as it comes near the view ----
  let scrollObserver = null;
  let ignoreScrollUntil = 0;

  function renderScrollList() {
    const deck = state.deck;
    const stage = $('stage');
    const oldTop = stage.classList.contains('scrollView') ? stage.scrollTop : null;
    const oldH = stage.scrollHeight || 1;
    stage.classList.add('scrollView');
    if (scrollObserver) scrollObserver.disconnect();
    scrollObserver = new IntersectionObserver((entries) => {
      for (const e of entries) if (e.isIntersecting) fillScrollItem(e.target);
    }, { root: stage, rootMargin: '800px 0px' });
    const list = el('div', 'scrollList');
    const w = Math.floor(deck.widthPx * state.scale);
    const h = Math.floor(deck.heightPx * state.scale);
    for (const s of deck.slides) {
      const item = el('div', 'scrollItem' + (s.hidden ? ' hiddenSlide' : '') + (s.index === state.current ? ' current' : ''));
      item.dataset.index = String(s.index);
      const head = el('div', 'scrollHead', 'Slide ' + s.index);
      if (s.hidden) head.appendChild(el('span', 'badge inline', 'Hidden slide'));
      item.appendChild(head);
      const ph = el('div', 'scrollPh');
      if (state.drawn) { ph.style.width = w + 'px'; ph.style.height = h + 'px'; }
      item.appendChild(ph);
      list.appendChild(item);
      scrollObserver.observe(item);
    }
    stage.replaceChildren(list);
    // Keep the place: after a zoom keep the same share of the column, else show the current slide.
    if (oldTop !== null && state.keepScroll) stage.scrollTop = oldTop * (stage.scrollHeight / oldH);
    else scrollToSlide(state.current);
    state.keepScroll = false;
  }

  function fillScrollItem(item) {
    if (item.dataset.filled) return;
    const i = parseInt(item.dataset.index, 10);
    const slide = state.deck.slides[i - 1];
    const node = slideClone(i, state.scale) || (state.drawn ? null : textSlide(slide));
    if (!node) return;
    item.querySelector('.scrollPh').replaceWith(node);
    item.dataset.filled = '1';
  }

  function scrollToSlide(n) {
    const stage = $('stage');
    const item = stage.querySelector('.scrollItem[data-index="' + n + '"]');
    if (!item) return;
    ignoreScrollUntil = performance.now() + 250;
    stage.scrollTop = item.offsetTop - 8;
  }

  let scrollRaf = 0;
  $('stage').addEventListener('scroll', () => {
    if (state.view !== 'scroll' || scrollRaf) return;
    scrollRaf = requestAnimationFrame(() => {
      scrollRaf = 0;
      if (performance.now() < ignoreScrollUntil || !state.deck) return;
      const stage = $('stage');
      const items = stage.querySelectorAll('.scrollItem');
      const tops = Array.from(items, (it) => it.offsetTop);
      let n = C.slideAtScroll(tops, stage.scrollTop, stage.clientHeight);
      if (stage.scrollTop + stage.clientHeight >= stage.scrollHeight - 2) n = state.deck.slides.length;
      if (n !== state.current) setCurrent(n, false);
    });
  });

  // ---- zoom and view ----
  function updateZoomControls() {
    const sel = $('zoomSelect');
    const custom = sel.querySelector('option[value="custom"]');
    const label = C.zoomLabel(state.zoom, state.scale);
    for (const o of sel.options) {
      if (o.value === 'fit') o.textContent = state.zoom === 'fit' ? label : 'Fit slide';
      if (o.value === 'width') o.textContent = state.zoom === 'width' ? label : 'Fit width';
    }
    if (typeof state.zoom === 'number') {
      const preset = Array.from(sel.options).find((o) => o.value !== 'custom' && Number(o.value) === state.zoom);
      if (preset) { custom.hidden = true; sel.value = preset.value; }
      else { custom.hidden = false; custom.textContent = label; sel.value = 'custom'; }
    } else { custom.hidden = true; sel.value = state.zoom; }
    $('zoomOutBtn').disabled = state.scale <= C.ZOOM_MIN + 0.001;
    $('zoomInBtn').disabled = state.scale >= C.ZOOM_MAX - 0.001;
  }

  function setZoom(z, anchor) {
    const before = state.scale;
    state.zoom = C.normalizeZoom(z);
    persistView();
    if (!state.deck) { updateZoomControls(); return; }
    const stage = $('stage');
    const left = stage.scrollLeft;
    const top = stage.scrollTop;
    state.keepScroll = true;
    renderStage();
    refreshThumbsLater();
    // Keep the point under the mouse (or the middle) where it was.
    const ratio = state.scale / (before || 1);
    const ax = anchor ? anchor.x : stage.clientWidth / 2;
    const ay = anchor ? anchor.y : stage.clientHeight / 2;
    if (state.view === 'slide') { stage.scrollLeft = (left + ax) * ratio - ax; stage.scrollTop = (top + ay) * ratio - ay; }
    else stage.scrollLeft = (left + ax) * ratio - ax;
  }

  function zoomStep(dir) { setZoom(C.stepZoom(state.scale, dir)); }

  function setView(v) {
    v = C.normalizeView(v);
    if (v === state.view) return;
    state.view = v;
    persistView();
    applyViewButtons();
    renderStage();
  }

  function applyViewButtons() {
    const on = state.view === 'scroll';
    $('scrollBtn').classList.toggle('on', on);
    $('scrollBtn').setAttribute('aria-pressed', String(on));
    document.body.classList.toggle('scrollMode', on);
  }

  function persistView() {
    saveState();
    vscode.postMessage({ type: 'viewPrefs', view: state.view, zoom: state.zoom });
  }

  let thumbTimer = null;
  function refreshThumbsLater() { clearTimeout(thumbTimer); thumbTimer = setTimeout(refreshThumbs, 50); }

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
    const clone = slideClone(parseInt(item.dataset.index, 10), C.fitScale(pic.clientWidth || 160, 1000, state.deck.widthPx, state.deck.heightPx));
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
    setCurrent(n, changed || force);
    if (state.view === 'scroll') {
      if (!$('stage').classList.contains('scrollView')) renderStage();
      scrollToSlide(n);
    }
  }

  // full: redraw the stage (one-slide view). The scroll view only refreshes the notes and the notice.
  function setCurrent(n, announce) {
    state.current = n;
    $('slideInput').value = String(n);
    for (const t of $('thumbs').children) {
      const on = t.dataset.index === String(n);
      t.classList.toggle('current', on);
      if (on) t.scrollIntoView({ block: 'nearest' });
    }
    for (const a of $('outline').querySelectorAll('.outlineItem')) a.classList.toggle('current', a.querySelector('.num').textContent === String(n));
    for (const it of $('stage').querySelectorAll('.scrollItem')) it.classList.toggle('current', it.dataset.index === String(n));
    if (state.view === 'scroll' && $('stage').classList.contains('scrollView')) renderSlideInfo();
    else renderStage();
    saveState();
    if (announce !== false) vscode.postMessage({ type: 'slideChanged', slide: n });
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
  $('zoomInBtn').addEventListener('click', () => zoomStep(1));
  $('zoomOutBtn').addEventListener('click', () => zoomStep(-1));
  $('zoomSelect').addEventListener('change', () => {
    const v = $('zoomSelect').value;
    if (v === 'custom') return;
    setZoom(v === 'fit' || v === 'width' ? v : Number(v));
  });
  $('scrollBtn').addEventListener('click', () => setView(state.view === 'scroll' ? 'slide' : 'scroll'));
  // Ctrl or Cmd + wheel, and a trackpad pinch, zoom around the mouse.
  $('stageWrap').addEventListener('wheel', (e) => {
    if (!(e.ctrlKey || e.metaKey) || !state.deck) return;
    e.preventDefault();
    const r = $('stage').getBoundingClientRect();
    setZoom(C.wheelZoom(state.scale, e.deltaMode === 1 ? e.deltaY * 16 : e.deltaY), { x: e.clientX - r.left, y: e.clientY - r.top });
  }, { passive: false });
  let searchTimer = null;
  $('searchInput').addEventListener('input', () => { clearTimeout(searchTimer); searchTimer = setTimeout(runSearch, 120); });
  $('searchInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); clearTimeout(searchTimer); if (state.query !== $('searchInput').value.trim()) runSearch(); nextHit(e.shiftKey ? -1 : 1); }
    else if (e.key === 'Escape') { $('searchInput').value = ''; runSearch(); setTab('slides'); }
  });
  document.addEventListener('keydown', (e) => {
    if (e.target instanceof HTMLInputElement || e.target instanceof HTMLSelectElement) return;
    if (!e.ctrlKey && !e.metaKey && !e.altKey) {
      if (e.key === '+' || e.key === '=') { e.preventDefault(); zoomStep(1); return; }
      if (e.key === '-' || e.key === '_') { e.preventDefault(); zoomStep(-1); return; }
      if (e.key === '0') { e.preventDefault(); setZoom('fit'); return; }
    }
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
