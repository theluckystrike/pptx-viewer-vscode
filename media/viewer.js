'use strict';
// Webview UI. All text from the deck is inserted with textContent or text nodes, never as HTML.
// Slide drawing is done by pptx-preview into a hidden host; we clone its slide nodes.
(function () {
  const vscode = acquireVsCodeApi();
  // Some shapes in real decks give the renderer a path with NaN coordinates. Such a path never draws;
  // write it as an empty path so the browser does not log an error for each one.
  const setAttr = Element.prototype.setAttribute;
  Element.prototype.setAttribute = function (name, value) {
    if (name === 'd' && typeof value === 'string' && value.includes('NaN')) value = 'M0,0';
    return setAttr.call(this, name, value);
  };
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
    drawn: null, drawError: '', config: { showNotes: true, renderSlides: true }, notes: true, presenting: false,
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
      if (m.pro && typeof m.pro === 'object') applyPro(m.pro);
    } else if (m.type === 'pro') {
      if (m.pro && typeof m.pro === 'object') applyPro(m.pro);
    } else if (m.type === 'exportSlides') {
      exportSlides(m.id, m.format === 'jpeg' ? 'jpeg' : 'png');
    } else if (m.type === 'exportCancel') {
      if (exportJob && exportJob.id === m.id) exportJob.cancelled = true;
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
      else if (m.name === 'present') setPresent(!state.presenting);
      else if (m.name === 'copySlidePng') copyPng();
      else if (m.name === 'saveSlidePng') savePng();
    }
  });

  function openExternalButton() {
    const b = el('button', null, 'Open in Default App');
    b.addEventListener('click', () => vscode.postMessage({ type: 'openInDefaultApp' }));
    return b;
  }

  // The toolbar works only when a deck with slides is open. Before that, and when a file cannot be
  // opened, the page box reads "– / –" and every deck button is greyed out (VER5 finding 2).
  const DECK_CONTROLS = ['prevBtn', 'nextBtn', 'slideInput', 'zoomOutBtn', 'zoomSelect', 'zoomInBtn', 'scrollBtn', 'presentBtn', 'searchInput', 'notesBtn', 'mdBtn', 'pngBtn', 'mdAllBtn'];
  function setDeckControls(on) {
    for (const id of DECK_CONTROLS) $(id).disabled = !on;
    document.body.classList.toggle('noDeck', !on);
    if (!on) { $('slideInput').value = '–'; $('slideCount').textContent = '/ –'; }
  }

  function showError(text) {
    if (state.presenting) setPresent(false);
    // Nothing from an earlier copy of the file stays on screen under the message.
    state.deck = null; state.drawn = null; state.renderer = null; drawGen++;
    $('stage').replaceChildren(); $('thumbs').replaceChildren(); $('outline').replaceChildren(); $('notes').replaceChildren();
    setDeckControls(false);
    const box = $('notice');
    box.replaceChildren(el('strong', null, 'Cannot open this file. '), document.createTextNode(text + ' '), openExternalButton());
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
    setDeckControls(true);
    buildThumbs();
    buildOutline();
    status(state.config.renderSlides && !deck.legacy && !deck.textOnly ? `Drawing ${n} slides…` : '');
    if (deck.textOnly) note(String(deck.textOnly));
    if (deck.hasMacros) note('This deck contains macros. They are never run.');
    if (deck.legacy) {
      // Old binary file: text only. Say so up front, with a way out to PowerPoint or Keynote.
      const s = $('status');
      s.replaceChildren(document.createTextNode('Old PowerPoint 97-2003 file. Text only: titles, slide text and speaker notes. Pictures, shapes and layout are not shown. '), openExternalButton());
      s.hidden = false;
    }
    go(state.current, true);
  }

  function note(text) {
    const s = $('status');
    s.textContent = text;
    s.hidden = false;
  }

  // ---- drawing with pptx-preview ----
  // The deck is read once; each slide is drawn the first time it is shown (stage, thumbnail, scroll
  // view, Present). A large deck shows its first slide without waiting for all the others, and a
  // slide the renderer cannot draw falls back to its text alone instead of stopping the whole deck.
  let drawGen = 0;
  async function draw(data) {
    const deck = state.deck;
    const gen = ++drawGen;
    if (!deck || typeof pptxPreview === 'undefined') { state.drawError = 'The slide renderer did not load.'; renderStage(); return; }
    const t0 = performance.now();
    const host = $('renderHost');
    host.replaceChildren();
    const box = el('div');
    host.appendChild(box);
    state.renderer = null;
    state.drawn = null;
    try {
      const buf = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
      const previewer = pptxPreview.init(box, { width: deck.widthPx, height: deck.heightPx, mode: 'list' });
      await previewer.load(buf);
      if (gen !== drawGen) return; // a newer copy of the file arrived meanwhile
      const count = previewer.pptx && previewer.pptx.slides ? previewer.pptx.slides.length : 0;
      state.renderer = { html: previewer.htmlRender, box, count };
      state.drawn = new Array(count).fill(undefined);
      status(deck.hasMacros ? 'This deck contains macros. They are never run.' : '');
      // Draw the slide on screen now, so the time reported is the time to the first picture.
      drawnNode(state.current);
      vscode.postMessage({ type: 'rendered', slides: count, ms: Math.round(performance.now() - t0) });
    } catch (e) {
      if (gen !== drawGen) return;
      state.drawn = null;
      state.renderer = null;
      state.drawError = 'Could not draw the slides (' + (e && e.message ? e.message : String(e)) + '). Text, notes and search still work.';
      status('');
      vscode.postMessage({ type: 'error', message: state.drawError });
    }
    refreshThumbs();
    renderStage();
    renderPresent();
  }

  // The drawn node of slide i (1-based), drawn now if needed. null when it cannot be drawn.
  function drawnNode(i) {
    const r = state.renderer;
    if (!r || !state.drawn || i < 1 || i > r.count) return null;
    const have = state.drawn[i - 1];
    if (have !== undefined) return have;
    let node = null;
    try {
      r.html.renderSlide(i - 1);
      const all = r.box.querySelectorAll('.pptx-preview-slide-wrapper-' + (i - 1));
      node = all.length ? all[all.length - 1] : null;
    } catch (e) {
      node = null;
      // Remove a half-drawn slide, if the renderer left one.
      for (const n of r.box.querySelectorAll('.pptx-preview-slide-wrapper-' + (i - 1))) n.remove();
      state.slideErrors = (state.slideErrors || 0) + 1;
    }
    state.drawn[i - 1] = node;
    return node;
  }

  function slideClone(i, scale) {
    const src = drawnNode(i);
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
    // The notice above the slide takes room; show it first, so Fit leaves room for it.
    renderSlideInfo();
    state.scale = currentScale();
    updateZoomControls();
    if (state.view === 'scroll') renderScrollList();
    else {
      const clone = slideClone(state.current, state.scale);
      stage.classList.remove('scrollView');
      stage.replaceChildren(clone || textSlide(slide));
      if (slide.hidden) stage.appendChild(el('div', 'badge', 'Hidden slide'));
    }
  }

  // Notice and notes for the current slide. Used on its own while the scroll view scrolls.
  function renderSlideInfo() {
    const slide = state.deck.slides[state.current - 1];
    const undrawable = state.view !== 'scroll' && state.drawn && state.drawn[state.current - 1] === null;
    const msg = state.drawError || (undrawable ? C.UNDRAWABLE_MESSAGE + ' ' : '') + C.noticeMessage(slide.notices);
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
    if (item.dataset.filled || !state.deck) return;
    const i = parseInt(item.dataset.index, 10);
    const slide = state.deck.slides[i - 1];
    const node = slideClone(i, state.scale) || textSlide(slide);
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
    if (!state.deck) return;
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
  $('exportBtn').addEventListener('click', () => vscode.postMessage({ type: 'exportMenu' }));
  $('pngBtn').addEventListener('click', (e) => { if (e.shiftKey) savePng(); else copyPng(); });
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
  // ---- slide as PNG: the drawn slide, about 1920 px wide, on white ----
  // The slide's own nodes are put in an SVG foreignObject (pictures are already data: URLs, so
  // nothing is fetched) and drawn on a canvas. Chart canvases are swapped for pictures of themselves.
  async function slidePng(n, jpeg) {
    const deck = state.deck;
    const src = deck && drawnNode(n);
    if (!src) throw new Error(state.drawn ? 'This slide could not be drawn, so there is no picture of it. Copy as Markdown copies its text.' : 'Old .ppt files and text-only mode have no slide pictures. Copy as Markdown copies the text.');
    const k = Math.min(3, Math.max(1, 1920 / deck.widthPx));
    const w = Math.round(deck.widthPx * k);
    const h = Math.round(deck.heightPx * k);
    const clone = src.cloneNode(true);
    const from = src.querySelectorAll('canvas');
    clone.querySelectorAll('canvas').forEach((c, i) => {
      try {
        const img = document.createElement('img');
        img.src = from[i].toDataURL('image/png');
        img.style.cssText = c.style.cssText;
        img.style.width = img.style.width || c.width + 'px';
        img.style.height = img.style.height || c.height + 'px';
        c.replaceWith(img);
      } catch (e) { c.remove(); }
    });
    Object.assign(clone.style, { position: 'absolute', left: '0', top: '0', margin: '0', width: deck.widthPx + 'px', height: deck.heightPx + 'px', transform: `scale(${k})`, transformOrigin: '0 0' });
    const xhtml = new XMLSerializer().serializeToString(clone);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><foreignObject x="0" y="0" width="${w}" height="${h}"><div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:${w}px;height:${h}px;overflow:hidden;background:#fff">${xhtml}</div></foreignObject></svg>`;
    const img = new Image();
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    await img.decode();
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(img, 0, 0);
    const blob = await new Promise((res) => (jpeg ? canvas.toBlob(res, 'image/jpeg', 0.92) : canvas.toBlob(res, 'image/png')));
    if (!blob) throw new Error('The picture could not be made.');
    return blob;
  }

  // ---- Zovo Pro: every slide as a picture, one message per slide (the host writes PNG files or builds a PDF) ----
  let exportJob = null;
  function applyPro(p) {
    $('exportGroup').hidden = p.show !== true;
    $('exportBadge').hidden = p.badge !== true;
  }
  // A MessageChannel turn, not setTimeout: a webview whose window is in the background gets its timers slowed to
  // one per second, which made a 120-slide export take two minutes.
  function yieldNow() {
    return new Promise((r) => { const ch = new MessageChannel(); ch.port1.onmessage = () => { ch.port1.close(); r(); }; ch.port2.postMessage(0); });
  }
  async function exportSlides(id, format) {
    const total = state.renderer ? state.renderer.count : 0;
    if (!state.deck || !total) { vscode.postMessage({ type: 'exportDone', id, error: 'This deck has no drawn slides to export.' }); return; }
    if (exportJob) exportJob.cancelled = true;
    const job = { id, cancelled: false };
    exportJob = job;
    const failed = [];
    for (let n = 1; n <= total; n++) {
      if (job.cancelled) break;
      status('Exporting slide ' + n + ' of ' + total + '…');
      try {
        const blob = await slidePng(n, format === 'jpeg');
        vscode.postMessage({ type: 'exportSlide', id, slide: n, data: new Uint8Array(await blob.arrayBuffer()) });
      } catch (e) {
        failed.push(n);
      }
      await yieldNow(); // keep the page responsive between slides
    }
    if (exportJob === job) exportJob = null;
    status('');
    vscode.postMessage({ type: 'exportDone', id, failed, cancelled: job.cancelled });
  }

  function pngFailed(e) {
    vscode.postMessage({ type: 'pngFailed', message: e && e.message ? e.message : String(e) });
  }

  async function copyPng() {
    if (!state.deck) return;
    const n = state.current;
    try {
      // The promise form keeps the click's permission to write the clipboard while the picture is made.
      const made = slidePng(n);
      made.catch(() => {});
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': made })]);
      vscode.postMessage({ type: 'pngDone', slide: n });
    } catch (e) {
      pngFailed(e && e.name === 'NotAllowedError' ? new Error('The clipboard did not accept the picture. Click in the slide and try again, or Shift+click Copy as PNG to save it as a file.') : e);
    }
  }

  async function savePng() {
    if (!state.deck) return;
    const n = state.current;
    try {
      const blob = await slidePng(n);
      vscode.postMessage({ type: 'savePng', slide: n, data: new Uint8Array(await blob.arrayBuffer()) });
    } catch (e) { pngFailed(e); }
  }

  // ---- Present: one slide at a time on black, fitted to the whole window ----
  function setPresent(on) {
    if (!state.deck || !state.deck.slides.length) return;
    if (on === state.presenting) return;
    state.presenting = on;
    $('present').hidden = !on;
    document.body.classList.toggle('presenting', on);
    vscode.postMessage({ type: 'present', on });
    if (on) { $('present').focus(); setTimeout(renderPresent, 60); setTimeout(renderPresent, 400); }
    else renderStage();
  }

  function renderPresent() {
    if (!state.presenting) return;
    const deck = state.deck;
    const box = $('presentStage');
    const fit = C.zoomScale('fit', Math.max(80, window.innerWidth - 8), Math.max(80, window.innerHeight - 28), deck.widthPx, deck.heightPx);
    const clone = slideClone(state.current, fit);
    if (clone) box.replaceChildren(clone);
    else { const t = textSlide(deck.slides[state.current - 1]); t.classList.add('presentText'); box.replaceChildren(t); }
    $('presentCount').textContent = `${state.current} / ${deck.slides.length}`;
  }

  $('presentBtn').addEventListener('click', () => setPresent(true));
  $('present').addEventListener('click', (e) => { if (e.target.closest && e.target.closest('#presentBar')) return; go(state.current + (e.shiftKey ? -1 : 1)); renderPresent(); });

  document.addEventListener('keydown', (e) => {
    if (state.presenting) {
      e.preventDefault();
      if (e.key === 'Escape') setPresent(false);
      else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(e.key)) { go(state.current + 1); renderPresent(); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(e.key)) { go(state.current - 1); renderPresent(); }
      else if (e.key === 'Home') { go(1); renderPresent(); }
      else if (e.key === 'End') { go(state.deck.slides.length); renderPresent(); }
      return;
    }
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
  window.addEventListener('resize', () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { if (state.presenting) renderPresent(); else renderStage(); }, 100); });

  setDeckControls(false);
  vscode.postMessage({ type: 'ready' });
})();
