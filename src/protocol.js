'use strict';
// Messages from the webview are untrusted. Accept only known shapes with checked values.

const MAX_SLIDE = 100000;
const VIEWS = new Set(['slide', 'scroll']);
const ZOOM_MIN = 0.1;
const ZOOM_MAX = 4;
const MAX_PNG_BYTES = 64 * 1024 * 1024;
const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function isPng(b) {
  return b.length > PNG_SIGNATURE.length && PNG_SIGNATURE.every((x, i) => b[i] === x);
}

/** 'fit', 'width' or a scale from 0.1 to 4 (1 = 100 %). Anything else is null. */
function zoomValue(v) {
  if (v === 'fit' || v === 'width') return v;
  return typeof v === 'number' && Number.isFinite(v) && v >= ZOOM_MIN && v <= ZOOM_MAX ? Math.round(v * 1000) / 1000 : null;
}

function int(v, lo, hi) {
  return Number.isInteger(v) && v >= lo && v <= hi ? v : null;
}

function parseWebviewMessage(raw) {
  if (!raw || typeof raw !== 'object' || typeof raw.type !== 'string') return null;
  switch (raw.type) {
    case 'ready':
    case 'loadAnyway':
    case 'copyDeckMarkdown':
    case 'openInDefaultApp':
      return { type: raw.type };
    case 'present':
      return typeof raw.on === 'boolean' ? { type: 'present', on: raw.on } : null;
    case 'copySlideMarkdown': {
      const slide = int(raw.slide, 1, MAX_SLIDE);
      return slide ? { type: raw.type, slide } : null;
    }
    case 'slideChanged': {
      const slide = int(raw.slide, 1, MAX_SLIDE);
      return slide ? { type: raw.type, slide } : null;
    }
    case 'rendered': {
      const slides = int(raw.slides, 0, MAX_SLIDE);
      const ms = int(raw.ms, 0, 3600000);
      return slides === null ? null : { type: raw.type, slides, ms: ms || 0 };
    }
    case 'viewPrefs': {
      const zoom = zoomValue(raw.zoom);
      if (!VIEWS.has(raw.view) || zoom === null) return null;
      return { type: raw.type, view: raw.view, zoom };
    }
    case 'savePng': {
      const slide = int(raw.slide, 1, MAX_SLIDE);
      const data = raw.data instanceof Uint8Array ? raw.data : raw.data instanceof ArrayBuffer ? new Uint8Array(raw.data) : null;
      return slide && data && isPng(data) && data.length <= MAX_PNG_BYTES ? { type: raw.type, slide, data } : null;
    }
    case 'pngDone': {
      const slide = int(raw.slide, 1, MAX_SLIDE);
      return slide ? { type: raw.type, slide } : null;
    }
    case 'exportMenu':
      return { type: raw.type };
    case 'exportSlide': {
      // One slide picture for a Zovo Pro export (PNG or JPEG; the host checks which one it asked for).
      const id = int(raw.id, 1, 1e9);
      const slide = int(raw.slide, 1, MAX_SLIDE);
      const data = raw.data instanceof Uint8Array ? raw.data : raw.data instanceof ArrayBuffer ? new Uint8Array(raw.data) : null;
      return id && slide && data && data.length > 8 && data.length <= MAX_PNG_BYTES ? { type: raw.type, id, slide, data } : null;
    }
    case 'exportDone': {
      const id = int(raw.id, 1, 1e9);
      if (!id) return null;
      const failed = Array.isArray(raw.failed) ? raw.failed.filter((n) => int(n, 1, MAX_SLIDE)).slice(0, 1000) : [];
      const out = { type: raw.type, id, failed, cancelled: raw.cancelled === true };
      if (typeof raw.error === 'string') out.error = raw.error.slice(0, 300);
      return out;
    }
    case 'pngFailed':
      return { type: raw.type, message: typeof raw.message === 'string' ? raw.message.slice(0, 300) : 'Unknown error' };
    case 'error':
      return { type: 'error', message: typeof raw.message === 'string' ? raw.message.slice(0, 500) : 'Unknown error' };
    default:
      return null;
  }
}

function readConfig(get) {
  const mb = Number(get('maxFileSizeMB', 300));
  return {
    maxFileSizeBytes: (Number.isFinite(mb) && mb > 0 ? mb : 300) * 1024 * 1024,
    showNotes: get('showNotes', true) !== false,
    renderSlides: get('renderSlides', true) !== false,
    defaultView: get('defaultView', 'slide') === 'scroll' ? 'scroll' : 'slide',
    defaultZoom: get('defaultZoom', 'fit') === 'width' ? 'width' : 'fit',
    presentFullScreen: get('presentFullScreen', true) !== false,
  };
}

module.exports = { parseWebviewMessage, readConfig, zoomValue, isPng };
