'use strict';
// Messages from the webview are untrusted. Accept only known shapes with checked values.

const MAX_SLIDE = 100000;
const VIEWS = new Set(['slide', 'scroll']);
const ZOOM_MIN = 0.1;
const ZOOM_MAX = 4;

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
      return { type: raw.type };
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
    case 'error':
      return { type: 'error', message: typeof raw.message === 'string' ? raw.message.slice(0, 500) : 'Unknown error' };
    default:
      return null;
  }
}

function readConfig(get) {
  const mb = Number(get('maxFileSizeMB', 150));
  return {
    maxFileSizeBytes: (Number.isFinite(mb) && mb > 0 ? mb : 150) * 1024 * 1024,
    showNotes: get('showNotes', true) !== false,
    renderSlides: get('renderSlides', true) !== false,
    defaultView: get('defaultView', 'slide') === 'scroll' ? 'scroll' : 'slide',
    defaultZoom: get('defaultZoom', 'fit') === 'width' ? 'width' : 'fit',
  };
}

module.exports = { parseWebviewMessage, readConfig, zoomValue };
