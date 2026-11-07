'use strict';
// Extract the text of a PowerPoint deck: slide order, titles, bullets with levels, tables,
// speaker notes, and the things we cannot draw (SmartArt, charts, video, audio, embedded objects).
const { readXmlParts, resolveTarget, relsPathFor } = require('./ooxml/zip');
const { parseXml, child, find, findAll, elements, local } = require('./ooxml/xml');

const REL_NOTES = /\/notesSlide$/;
const REL_DIAGRAM_DRAWING = /\/diagramDrawing$/;
const SKIP_PH = new Set(['sldNum', 'dt', 'ftr', 'hdr', 'sldImg']);
const TITLE_PH = new Set(['title', 'ctrTitle']);
const EMU_PER_PX = 9525; // 914400 EMU per inch / 96 px per inch

function parseRels(xml) {
  const map = new Map();
  if (!xml) return map;
  for (const r of findAll(parseXml(xml), 'Relationship')) {
    map.set(r.attrs.Id, { type: r.attrs.Type || '', target: r.attrs.Target || '', external: r.attrs.TargetMode === 'External' });
  }
  return map;
}

function placeholderType(sp) {
  const nv = elements(sp).find((e) => local(e.name).startsWith('nv'));
  const ph = nv && find(nv, 'p:ph');
  if (!ph) return null;
  return ph.attrs.type || 'body';
}

function paragraphText(p) {
  let s = '';
  for (const c of elements(p)) {
    const n = local(c.name);
    if (n === 'r' || n === 'fld') {
      const t = child(c, 'a:t');
      if (t) s += t.children.filter((x) => typeof x === 'string').join('');
    } else if (n === 'br') s += '\n';
  }
  return s.replace(/\u000b/g, '\n');
}

function textBodyParagraphs(txBody) {
  const out = [];
  for (const p of elements(txBody)) {
    if (p.name !== 'a:p') continue;
    const text = paragraphText(p).trim();
    if (!text) continue;
    const pPr = child(p, 'a:pPr');
    const level = pPr && pPr.attrs.lvl ? Math.min(8, parseInt(pPr.attrs.lvl, 10) || 0) : 0;
    out.push({ text, level });
  }
  return out;
}

function tableRows(tbl) {
  return elements(tbl).filter((e) => e.name === 'a:tr').map((tr) =>
    elements(tr).filter((e) => e.name === 'a:tc').map((tc) => {
      const body = child(tc, 'a:txBody');
      return body ? textBodyParagraphs(body).map((p) => p.text).join(' ') : '';
    }));
}

/** Walk a shape tree in document order. */
function walkShapes(tree, slide) {
  for (const el of elements(tree)) {
    const n = el.name;
    if (n === 'p:sp') {
      const ph = placeholderType(el);
      if (ph && SKIP_PH.has(ph)) continue;
      const body = child(el, 'p:txBody');
      if (!body) continue;
      const paras = textBodyParagraphs(body);
      if (!paras.length) continue;
      if (ph && TITLE_PH.has(ph) && !slide.title) slide.title = paras.map((p) => p.text).join(' ');
      else slide.blocks.push({ kind: 'text', subtitle: ph === 'subTitle', paragraphs: paras });
    } else if (n === 'p:grpSp') {
      walkShapes(el, slide);
    } else if (n === 'p:graphicFrame') {
      const gd = find(el, 'a:graphicData');
      const uri = gd ? gd.attrs.uri || '' : '';
      const tbl = gd && find(gd, 'a:tbl');
      if (tbl) slide.blocks.push({ kind: 'table', rows: tableRows(tbl) });
      else if (/\/diagram$/.test(uri)) slide.diagrams = (slide.diagrams || 0) + 1;
      else if (/\/chart$/.test(uri) || /chartex/i.test(uri)) addNotice(slide, 'chart');
      else if (/\/ole$/.test(uri)) addNotice(slide, 'object');
    } else if (n === 'p:pic') {
      const nvPr = find(el, 'p:nvPr');
      if (nvPr && (find(nvPr, 'a:videoFile') || find(nvPr, 'a:quickTimeFile'))) addNotice(slide, 'video');
      else if (nvPr && (find(nvPr, 'a:audioFile') || find(nvPr, 'a:wavAudioFile'))) addNotice(slide, 'audio');
      else if (nvPr && find(nvPr, 'p14:media')) addNotice(slide, 'video');
      else slide.images += 1;
    } else if (n === 'mc:AlternateContent') {
      // Prefer the Choice branch (newer content); the Fallback is usually a picture of it.
      const choice = child(el, 'mc:Choice') || child(el, 'mc:Fallback');
      if (choice) walkShapes(choice, slide);
    } else if (n === 'p:contentPart') {
      addNotice(slide, 'ink');
    }
  }
}

function addNotice(slide, kind) {
  if (!slide.notices.includes(kind)) slide.notices.push(kind);
}

function notesText(xml) {
  const doc = parseXml(xml);
  const tree = find(doc, 'p:spTree');
  if (!tree) return '';
  const lines = [];
  for (const sp of findAll(tree, 'p:sp')) {
    const ph = placeholderType(sp);
    if (ph !== 'body') continue;
    const body = child(sp, 'p:txBody');
    if (body) for (const p of textBodyParagraphs(body)) lines.push(p.text);
  }
  return lines.join('\n');
}

const MAIN_TYPES = {
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml': 'pptx',
  'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml': 'pptm',
  'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml': 'ppsx',
  'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml': 'ppsm',
  'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml': 'potx',
  'application/vnd.ms-powerpoint.template.macroEnabled.main+xml': 'potm',
};

/** Parse deck bytes into a plain, JSON-safe object. Never runs macros or scripts; it only reads XML. */
function parseDeck(bytes) {
  const parts = readXmlParts(bytes);
  const ct = parts.get('[Content_Types].xml');
  let mainPart = 'ppt/presentation.xml';
  let kind = 'pptx';
  if (ct) {
    for (const o of findAll(parseXml(ct), 'Override')) {
      const k = MAIN_TYPES[o.attrs.ContentType];
      if (k) { mainPart = (o.attrs.PartName || '').replace(/^\//, ''); kind = k; break; }
    }
  }
  const presXml = parts.get(mainPart);
  if (!presXml) throw new Error('No presentation part found. This is not a PowerPoint file.');
  const pres = parseXml(presXml);
  const rels = parseRels(parts.get(relsPathFor(mainPart)));
  const sz = find(pres, 'p:sldSz');
  const width = sz ? parseInt(sz.attrs.cx, 10) || 9144000 : 9144000;
  const height = sz ? parseInt(sz.attrs.cy, 10) || 6858000 : 6858000;
  const ids = find(pres, 'p:sldIdLst');
  const slides = [];
  const hasMacros = /\.bin$/i.test([...rels.values()].map((r) => r.target).join('\n')) || kind.endsWith('m');
  for (const sid of ids ? elements(ids).filter((e) => e.name === 'p:sldId') : []) {
    const rel = rels.get(sid.attrs['r:id']);
    if (!rel || rel.external) continue;
    const part = resolveTarget(mainPart, rel.target);
    const xml = parts.get(part);
    const slide = { index: slides.length + 1, title: '', blocks: [], notes: '', notices: [], images: 0, hidden: false };
    if (!xml) { slide.notices.push('missing'); slides.push(slide); continue; }
    const doc = parseXml(xml);
    const sld = find(doc, 'p:sld');
    if (sld && sld.attrs.show === '0') slide.hidden = true;
    const tree = find(doc, 'p:spTree');
    if (tree) walkShapes(tree, slide);
    const srels = parseRels(parts.get(relsPathFor(part)));
    // PowerPoint saves a drawn copy of each SmartArt (ppt/diagrams/drawingN.xml). The renderer draws
    // that copy, so only SmartArt without one (some other apps' files) gets the "cannot draw" notice.
    const drawings = [...srels.values()].filter((r) => REL_DIAGRAM_DRAWING.test(r.type) && !r.external && parts.has(resolveTarget(part, r.target))).length;
    if (slide.diagrams > drawings) addNotice(slide, 'smartart');
    delete slide.diagrams;
    for (const r of srels.values()) {
      if (REL_NOTES.test(r.type) && !r.external) {
        const nx = parts.get(resolveTarget(part, r.target));
        if (nx) slide.notes = notesText(nx);
      }
    }
    if (!slide.title) {
      const first = slide.blocks.find((b) => b.kind === 'text');
      if (first) slide.fallbackTitle = first.paragraphs[0].text.slice(0, 80);
    }
    slides.push(slide);
  }
  return {
    kind,
    hasMacros,
    width,
    height,
    widthPx: Math.round(width / EMU_PER_PX),
    heightPx: Math.round(height / EMU_PER_PX),
    slides,
  };
}

/** Plain text used by search: title, body, table cells and notes. */
function slideSearchText(slide) {
  const parts = [slide.title];
  for (const b of slide.blocks) {
    if (b.kind === 'text') for (const p of b.paragraphs) parts.push(p.text);
    else if (b.kind === 'table') for (const r of b.rows) parts.push(r.join(' '));
  }
  parts.push(slide.notes);
  return parts.filter(Boolean).join('\n');
}

function mdEscapeInline(s) {
  return s.replace(/\\/g, '\\\\').replace(/([*_`[\]#<>|])/g, '\\$1').replace(/\n/g, ' ');
}

function mdCell(s) {
  return mdEscapeInline(s).trim() || ' ';
}

const NOTICE_LABEL = {
  smartart: 'SmartArt', chart: 'chart', video: 'video', audio: 'audio', object: 'embedded object', ink: 'ink drawing', missing: 'missing slide part',
};

function slideToMarkdown(slide, opts = {}) {
  const lines = [];
  const title = slide.title || slide.fallbackTitle || `Slide ${slide.index}`;
  lines.push(`${opts.heading || '##'} ${mdEscapeInline(title)}`);
  for (const b of slide.blocks) {
    lines.push('');
    if (b.kind === 'text') {
      if (b.paragraphs.length === 1 && b.subtitle) lines.push(`*${mdEscapeInline(b.paragraphs[0].text)}*`);
      else for (const p of b.paragraphs) lines.push(`${'  '.repeat(p.level)}- ${mdEscapeInline(p.text)}`);
    } else if (b.kind === 'table' && b.rows.length) {
      const cols = Math.max(...b.rows.map((r) => r.length));
      const pad = (r) => r.concat(Array(cols - r.length).fill(''));
      lines.push(`| ${pad(b.rows[0]).map(mdCell).join(' | ')} |`);
      lines.push(`|${' --- |'.repeat(cols)}`);
      for (const r of b.rows.slice(1)) lines.push(`| ${pad(r).map(mdCell).join(' | ')} |`);
    }
  }
  if (slide.notices.length) {
    lines.push('', `*Not shown as text: ${slide.notices.map((n) => NOTICE_LABEL[n] || n).join(', ')}.*`);
  }
  if (slide.notes && opts.notes !== false) {
    lines.push('', '> **Notes:** ' + slide.notes.split('\n').map(mdEscapeInline).join('\n> '));
  }
  return lines.join('\n') + '\n';
}

function deckToMarkdown(deck, fileName) {
  const head = fileName ? `# ${mdEscapeInline(fileName)}\n\n` : '';
  return head + deck.slides.map((s) => slideToMarkdown(s)).join('\n');
}

module.exports = { parseDeck, slideToMarkdown, deckToMarkdown, slideSearchText, mdEscapeInline, NOTICE_LABEL };
