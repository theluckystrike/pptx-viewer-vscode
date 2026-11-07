'use strict';
// Text of an old binary PowerPoint 97-2003 file (.ppt, .pps, .pot): slide titles, body text and
// speaker notes, in slide order. Read-only and offline, no LibreOffice. Pictures, shapes and layout
// are not read. Formats: [MS-CFB] (the OLE container) and [MS-PPT] (the records inside it).

const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const END = 0xfffffffe;
const FREE = 0xffffffff;

// [MS-PPT] record types used here
const RT = {
  Document: 0x03e8, Slide: 0x03ee, SlideAtom: 0x03ef, Notes: 0x03f0, NotesAtom: 0x03f1,
  SlidePersistAtom: 0x03f3, SlideListWithText: 0x0ff0, TextHeaderAtom: 0x0f9f,
  TextCharsAtom: 0x0fa0, TextBytesAtom: 0x0fa8, PersistDirectoryAtom: 0x1772,
  CryptSession10Container: 0x2f14,
};
// TextHeaderAtom textType: 0 title, 1 body, 2 notes, 4 other, 5 center body, 6 center title, 7 half body, 8 quarter body
const TITLE_TYPES = new Set([0, 6]);

function isOle(b) { return b && b.length >= 512 && OLE_MAGIC.every((v, i) => b[i] === v); }

const u16 = (b, o) => b[o] | (b[o + 1] << 8);
const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

/** Minimal [MS-CFB] reader. Returns a function name -> stream bytes (or null). Throws on a broken container. */
function openCfb(b) {
  if (!isOle(b)) throw new Error('not an OLE file');
  const shift = u16(b, 0x1e);
  if (shift !== 9 && shift !== 12) throw new Error('bad sector size');
  const ss = 1 << shift;
  const miniShift = u16(b, 0x20);
  const cutoff = u32(b, 0x38) || 4096;
  const sectors = Math.floor((b.length - ss) / ss) + 1;
  const off = (s) => (s + 1) * ss;
  // FAT sector list from the header DIFAT and the DIFAT chain
  const fatSecs = [];
  for (let i = 0; i < 109; i++) { const s = u32(b, 0x4c + i * 4); if (s < sectors) fatSecs.push(s); }
  let dif = u32(b, 0x44);
  for (let guard = 0; dif < sectors && guard < sectors; guard++) {
    const o = off(dif);
    for (let i = 0; i < ss / 4 - 1; i++) { const s = u32(b, o + i * 4); if (s < sectors) fatSecs.push(s); }
    dif = u32(b, o + ss - 4);
  }
  const per = ss / 4;
  const fat = (s) => {
    const fs = fatSecs[Math.floor(s / per)];
    if (fs === undefined) return END;
    const o = off(fs) + (s % per) * 4;
    return o + 4 <= b.length ? u32(b, o) : END;
  };
  const chain = (start, next, limit) => {
    const out = [];
    const seen = new Set();
    for (let s = start; s !== END && s !== FREE && out.length < limit; s = next(s)) {
      if (seen.has(s)) throw new Error('loop in the sector chain');
      seen.add(s);
      out.push(s);
    }
    return out;
  };
  const readChain = (start, size) => {
    const out = new Uint8Array(size);
    let p = 0;
    for (const s of chain(start, fat, sectors)) {
      const o = off(s);
      if (o >= b.length) break;
      const n = Math.min(ss, size - p, b.length - o);
      out.set(b.subarray(o, o + n), p);
      p += n;
      if (p >= size) break;
    }
    return p < size ? out.subarray(0, p) : out;
  };
  // Directory
  const dirSecs = chain(u32(b, 0x30), fat, sectors);
  const entries = [];
  for (const s of dirSecs) {
    for (let i = 0; i < ss / 128; i++) {
      const o = off(s) + i * 128;
      if (o + 128 > b.length) break;
      const nameLen = Math.min(64, u16(b, o + 64));
      let name = '';
      for (let k = 0; k + 1 < nameLen - 1; k += 2) name += String.fromCharCode(u16(b, o + k));
      entries.push({ name, type: b[o + 66], start: u32(b, o + 116), size: u32(b, o + 120) });
    }
  }
  const root = entries.find((e) => e.type === 5);
  if (!root) throw new Error('no root entry');
  let mini = null;
  let miniFat = null;
  const getMini = () => {
    if (!mini) {
      mini = readChain(root.start, root.size);
      const mfSecs = chain(u32(b, 0x3c), fat, sectors);
      miniFat = new Uint32Array(mfSecs.length * per);
      mfSecs.forEach((s, i) => { for (let k = 0; k < per; k++) miniFat[i * per + k] = u32(b, off(s) + k * 4); });
    }
  };
  return function stream(name) {
    const e = entries.find((x) => x.type === 2 && x.name === name);
    if (!e) return null;
    if (e.size >= cutoff) return readChain(e.start, e.size);
    getMini();
    const msz = 1 << miniShift;
    const out = new Uint8Array(e.size);
    let p = 0;
    for (const s of chain(e.start, (x) => (x < miniFat.length ? miniFat[x] : END), miniFat.length)) {
      const n = Math.min(msz, e.size - p);
      out.set(mini.subarray(s * msz, s * msz + n), p);
      p += n;
      if (p >= e.size) break;
    }
    return out;
  };
}

/** Records of one level: [{type, inst, ver, start, len, body}] where body is the offset of the data. */
function records(d, start, end) {
  const out = [];
  let p = start;
  while (p + 8 <= end) {
    const vi = u16(d, p);
    const len = u32(d, p + 4);
    const body = p + 8;
    if (body + len > end) break;
    out.push({ ver: vi & 0xf, inst: vi >> 4, type: u16(d, p + 2), start: p, len, body });
    p = body + len;
  }
  return out;
}

function atomText(d, r) {
  if (r.type === RT.TextCharsAtom) {
    let s = '';
    for (let i = 0; i + 1 < r.len; i += 2) s += String.fromCharCode(u16(d, r.body + i));
    return s;
  }
  let s = '';
  for (let i = 0; i < r.len; i++) s += String.fromCharCode(d[r.body + i]);
  return s;
}

/** Text runs inside a container, depth first: [{textType, text}]. A TextHeaderAtom sets the type of the next text atom. */
function textRuns(d, start, end, out = [], depth = 0) {
  if (depth > 32) return out;
  let type = 4;
  for (const r of records(d, start, end)) {
    if (r.ver === 0xf) textRuns(d, r.body, r.body + r.len, out, depth + 1);
    else if (r.type === RT.TextHeaderAtom && r.len >= 4) type = u32(d, r.body);
    else if (r.type === RT.TextCharsAtom || r.type === RT.TextBytesAtom) out.push({ textType: type, text: atomText(d, r) });
  }
  return out;
}

/** Paragraphs of one text run. \r ends a paragraph, \v is a line break. Placeholder prompts are left out by PowerPoint already. */
function paragraphs(text) {
  // A lone "*" is the slide number field of a placeholder, not text.
  return text.split('\r').map((t) => t.replace(/\u000b/g, '\n').replace(/[\u0000-\u0008\u000e-\u001f]/g, '').trim()).filter((t) => t && t !== '*');
}

/**
 * Parse a .ppt file into the same deck shape as parseDeck (pptx.js), with kind 'ppt' and legacy: true.
 * Throws an Error with a message for the user when the file cannot be read.
 */
function parsePpt(bytes) {
  let stream;
  try { stream = openCfb(bytes); } catch (e) { throw new Error('This old PowerPoint file is damaged and cannot be read (' + e.message + ').'); }
  const d = stream('PowerPoint Document');
  if (!d) {
    if (stream('EncryptedPackage')) throw new Error('This deck is protected with a password, so it cannot be read.');
    throw new Error('This is an OLE file but not a PowerPoint 97-2003 deck (no "PowerPoint Document" stream).');
  }
  const top = records(d, 0, d.length);
  // Persist directory: persist id -> offset of the record in the stream. Later edits win.
  const persist = new Map();
  for (const r of top) {
    if (r.type !== RT.PersistDirectoryAtom) continue;
    let p = r.body;
    const end = r.body + r.len;
    while (p + 4 <= end) {
      const v = u32(d, p);
      const id = v & 0xfffff;
      const n = v >>> 20;
      p += 4;
      for (let i = 0; i < n && p + 4 <= end; i++, p += 4) persist.set(id + i, u32(d, p));
    }
  }
  const byStart = new Map(top.map((r) => [r.start, r]));
  const docs = top.filter((r) => r.type === RT.Document);
  const doc = docs[docs.length - 1];
  if (!doc) throw new Error('This old PowerPoint file has no document record, so it cannot be read.');
  const docRecs = records(d, doc.body, doc.body + doc.len);
  if (docRecs.some((r) => r.type === RT.CryptSession10Container) || top.some((r) => r.type === RT.CryptSession10Container)) {
    throw new Error('This deck is protected with a password, so it cannot be read.');
  }
  // Notes by the slide id they belong to
  const notesBySlide = new Map();
  for (const r of top) {
    if (r.type !== RT.Notes) continue;
    const inner = records(d, r.body, r.body + r.len);
    const atom = inner.find((x) => x.type === RT.NotesAtom && x.len >= 4);
    if (!atom) continue;
    const text = textRuns(d, r.body, r.body + r.len).filter((t) => t.textType === 2 || t.textType === 1 || t.textType === 4)
      .flatMap((t) => paragraphs(t.text)).join('\n');
    if (text) notesBySlide.set(u32(d, atom.body), text);
  }
  // Slides in order, from the slide list (instance 0) of the document
  const list = docRecs.find((r) => r.type === RT.SlideListWithText && r.inst === 0);
  const slides = [];
  if (list) {
    let cur = null;
    for (const r of records(d, list.body, list.body + list.len)) {
      if (r.type === RT.SlidePersistAtom && r.len >= 16) {
        cur = { persistId: u32(d, r.body), slideId: u32(d, r.body + 12), runs: [] };
        slides.push(cur);
      } else if (cur && r.type === RT.TextHeaderAtom && r.len >= 4) {
        cur.runs.push({ textType: u32(d, r.body), text: '' });
      } else if (cur && (r.type === RT.TextCharsAtom || r.type === RT.TextBytesAtom)) {
        const last = cur.runs[cur.runs.length - 1];
        if (last && !last.text) last.text = atomText(d, r);
        else cur.runs.push({ textType: 4, text: atomText(d, r) });
      }
    }
  }
  const out = slides.map((s, i) => {
    // Text boxes that are not placeholders live in the slide record itself.
    const at = persist.get(s.persistId);
    const rec = at !== undefined ? byStart.get(at) : null;
    const runs = s.runs.slice();
    if (rec && rec.type === RT.Slide) {
      const seen = new Set(runs.map((r) => r.text));
      for (const t of textRuns(d, rec.body, rec.body + rec.len)) if (!seen.has(t.text)) { seen.add(t.text); runs.push(t); }
    }
    const slide = { index: i + 1, title: '', blocks: [], notes: notesBySlide.get(s.slideId) || '', notices: [], images: 0, hidden: false };
    for (const r of runs) {
      const paras = paragraphs(r.text);
      if (!paras.length) continue;
      if (TITLE_TYPES.has(r.textType) && !slide.title) slide.title = paras.join(' ');
      else slide.blocks.push({ kind: 'text', subtitle: r.textType === 5, paragraphs: paras.map((text) => ({ text, level: 0 })) });
    }
    if (!slide.title) {
      const first = slide.blocks[0];
      if (first) slide.fallbackTitle = first.paragraphs[0].text.slice(0, 80);
    }
    return slide;
  });
  return { kind: 'ppt', legacy: true, hasMacros: false, width: 9144000, height: 6858000, widthPx: 960, heightPx: 720, slides: out };
}

/** What a file with a .ppt name really is, from its first bytes: 'ole', 'zip' (a renamed .pptx) or 'other'. */
function sniffPpt(bytes) {
  if (isOle(bytes)) return 'ole';
  if (bytes && bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 3 && bytes[3] === 4) return 'zip';
  return 'other';
}

module.exports = { parsePpt, openCfb, sniffPpt, isOle };
