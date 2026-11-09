'use strict';
// A small PDF writer for "Export the deck as PDF" (Zovo Pro): one page per slide, each page one JPEG picture of the
// slide at the slide size. JPEG goes into a PDF as is (DCTDecode), so no image code is needed. The JPEG bytes come
// from the webview, so jpegInfo() checks them before they are used.

/** Width, height and color components of a baseline or progressive JPEG, or null when it is not one. */
function jpegInfo(b) {
  if (!b || b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return null;
  let i = 2;
  while (i + 4 <= b.length) {
    if (b[i] !== 0xff) return null;
    let m = b[i + 1];
    while (m === 0xff && i + 2 < b.length) { i++; m = b[i + 1]; } // fill bytes
    if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { i += 2; continue; }
    if (m === 0xd9 || m === 0xda) return null; // end or scan before a frame header
    const len = (b[i + 2] << 8) | b[i + 3];
    if (len < 2 || i + 2 + len > b.length) return null;
    const isSof = m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc;
    if (isSof) {
      if (len < 8) return null;
      const height = (b[i + 5] << 8) | b[i + 6];
      const width = (b[i + 7] << 8) | b[i + 8];
      const comps = b[i + 9];
      if (!width || !height || ![1, 3, 4].includes(comps)) return null;
      return { width, height, comps };
    }
    i += 2 + len;
  }
  return null;
}

// PDF text string: printable ASCII as a literal string, anything else as UTF-16BE hex.
function pdfString(s) {
  const t = String(s == null ? '' : s).slice(0, 300);
  if (/^[\x20-\x7e]*$/.test(t)) return '(' + t.replace(/[\\()]/g, (c) => '\\' + c) + ')';
  let hex = 'FEFF';
  for (const ch of t) {
    const cp = ch.codePointAt(0);
    const units = cp > 0xffff ? [0xd800 + ((cp - 0x10000) >> 10), 0xdc00 + ((cp - 0x10000) & 0x3ff)] : [cp];
    for (const u of units) hex += u.toString(16).toUpperCase().padStart(4, '0');
  }
  return '<' + hex + '>';
}

const num = (n) => (Math.round(n * 100) / 100).toString();

/**
 * pages: [{ jpeg: Uint8Array }], opts: { widthPt, heightPt, title, creator }. Every page gets the same size
 * (the slide size in points, 72 per inch). Returns the PDF as a Buffer. Throws on a bad picture.
 */
function imagesToPdf(pages, opts) {
  const o = opts || {};
  if (!Array.isArray(pages) || !pages.length) throw new Error('No slides to write.');
  const W = Number(o.widthPt) > 0 ? Number(o.widthPt) : 720;
  const H = Number(o.heightPt) > 0 ? Number(o.heightPt) : 405;
  const chunks = [];
  let size = 0;
  const offsets = [];
  const push = (x) => { const b = typeof x === 'string' ? Buffer.from(x, 'latin1') : Buffer.from(x.buffer, x.byteOffset, x.byteLength); chunks.push(b); size += b.length; };
  const obj = (n, body) => { offsets[n] = size; push(n + ' 0 obj\n'); body(); push('\nendobj\n'); };
  push('%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');
  // Objects: 1 catalog, 2 pages, 3 info, then 3 per page (page, content, image).
  const first = 4;
  const kids = pages.map((_, i) => (first + i * 3) + ' 0 R').join(' ');
  obj(1, () => push('<< /Type /Catalog /Pages 2 0 R >>'));
  obj(2, () => push('<< /Type /Pages /Count ' + pages.length + ' /Kids [' + kids + '] >>'));
  obj(3, () => push('<< /Title ' + pdfString(o.title || '') + ' /Creator ' + pdfString(o.creator || 'PPTX Viewer by Zovo') + ' /Producer ' + pdfString('PPTX Viewer by Zovo') + ' >>'));
  pages.forEach((p, i) => {
    const info = jpegInfo(p && p.jpeg);
    if (!info) throw new Error('Slide ' + (i + 1) + ' is not a JPEG picture.');
    const pn = first + i * 3;
    const content = 'q ' + num(W) + ' 0 0 ' + num(H) + ' 0 0 cm /S Do Q\n';
    obj(pn, () => push('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + num(W) + ' ' + num(H) + '] /Resources << /XObject << /S ' + (pn + 2) + ' 0 R >> >> /Contents ' + (pn + 1) + ' 0 R >>'));
    obj(pn + 1, () => { push('<< /Length ' + Buffer.byteLength(content, 'latin1') + ' >>\nstream\n'); push(content); push('\nendstream'); });
    const cs = info.comps === 1 ? '/DeviceGray' : info.comps === 4 ? '/DeviceCMYK /Decode [1 0 1 0 1 0 1 0]' : '/DeviceRGB';
    obj(pn + 2, () => {
      push('<< /Type /XObject /Subtype /Image /Width ' + info.width + ' /Height ' + info.height + ' /ColorSpace ' + cs + ' /BitsPerComponent 8 /Filter /DCTDecode /Length ' + p.jpeg.length + ' >>\nstream\n');
      push(p.jpeg);
      push('\nendstream');
    });
  });
  const count = first + pages.length * 3;
  const xref = size;
  let x = 'xref\n0 ' + count + '\n0000000000 65535 f \n';
  for (let n = 1; n < count; n++) x += String(offsets[n]).padStart(10, '0') + ' 00000 n \n';
  push(x);
  push('trailer\n<< /Size ' + count + ' /Root 1 0 R /Info 3 0 R >>\nstartxref\n' + xref + '\n%%EOF\n');
  return Buffer.concat(chunks, size);
}

/** Slide size in points from the deck size in CSS pixels (96 per inch). */
function slideSizePt(deck) {
  const w = deck && deck.widthPx > 0 ? deck.widthPx : 960;
  const h = deck && deck.heightPx > 0 ? deck.heightPx : 540;
  return { widthPt: w * 0.75, heightPt: h * 0.75 };
}

module.exports = { jpegInfo, imagesToPdf, pdfString, slideSizePt };
