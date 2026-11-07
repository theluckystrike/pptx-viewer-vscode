'use strict';
// Read only the XML parts of an OOXML package (pptx, docx, xlsx). Media is skipped,
// so a deck full of videos still opens fast. Shared design with the DOCX viewer.
const { unzipSync, strFromU8 } = require('fflate');

const MAX_PART_BYTES = 32 * 1024 * 1024; // one XML part
const MAX_TOTAL_BYTES = 256 * 1024 * 1024; // all XML parts together

class PackageError extends Error {}

function isZip(bytes) {
  return bytes && bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && (bytes[2] === 3 || bytes[2] === 5 || bytes[2] === 7);
}

/** Returns a Map of part name -> XML string for .xml and .rels parts. */
function readXmlParts(bytes) {
  if (!isZip(bytes)) throw new PackageError('This file is not a ZIP package. It may be an old binary .ppt file or damaged.');
  let total = 0;
  let maxPart = 0;
  let files;
  try {
    files = unzipSync(bytes, {
      filter(f) {
        if (!/\.(xml|rels)$/i.test(f.name)) return false;
        maxPart = Math.max(maxPart, f.originalSize);
        if (f.originalSize > MAX_PART_BYTES) return false;
        total += f.originalSize;
        if (total > MAX_TOTAL_BYTES) throw new PackageError('The XML inside this file is larger than 256 MB. It was not opened.');
        return true;
      },
    });
  } catch (e) {
    if (e instanceof PackageError) throw e;
    throw new PackageError('The file could not be unzipped: ' + (e && e.message ? e.message : String(e)));
  }
  const parts = new Map();
  let inflated = 0;
  for (const [name, data] of Object.entries(files)) {
    maxPart = Math.max(maxPart, data.length);
    inflated += data.length;
    if (data.length > MAX_PART_BYTES) continue; // header lied about the size
    parts.set(name.replace(/^\/+/, ''), strFromU8(data));
  }
  // Unpacked XML sizes (largest part, all parts), so a caller can refuse to hand a zip bomb to a renderer.
  parts.xmlStats = { maxPart, total: Math.max(total, inflated) };
  return parts;
}

/** Resolve a relationship target against the folder of the source part. */
function resolveTarget(sourcePart, target) {
  if (target.startsWith('/')) return target.slice(1);
  const base = sourcePart.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') base.pop();
    else if (seg && seg !== '.') base.push(seg);
  }
  return base.join('/');
}

function relsPathFor(part) {
  const i = part.lastIndexOf('/');
  return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels';
}

module.exports = { readXmlParts, resolveTarget, relsPathFor, isZip, PackageError };
