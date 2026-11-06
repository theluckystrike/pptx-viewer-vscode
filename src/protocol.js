'use strict';
// Messages from the webview are untrusted. Accept only known shapes with checked values.

const MAX_SLIDE = 100000;

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
  };
}

module.exports = { parseWebviewMessage, readConfig };
