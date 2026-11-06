'use strict';
const crypto = require('crypto');
const { footerHtml, FOOTER_CSS } = require('./support');

function escapeAttr(s) {
  return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Webview HTML. Scripts run only with the nonce, from the extension's media folder.
 * Images may be data: or blob: URLs made from the deck itself. No network: connect-src is none.
 * style-src needs 'unsafe-inline' because pptx-preview positions every shape with inline styles.
 */
function buildCsp(cspSource, nonce) {
  return [
    "default-src 'none'",
    `img-src ${cspSource} data: blob:`,
    `style-src ${cspSource} 'unsafe-inline'`,
    `font-src ${cspSource} data:`,
    `script-src 'nonce-${nonce}'`,
    'media-src blob:',
    "connect-src 'none'",
    "frame-src 'none'",
    "object-src 'none'",
    "base-uri 'none'",
    "form-action 'none'",
  ].join('; ');
}

// opts.showSupportLinks: false hides the small Sponsor and More free viewers footer (setting pptxViewer.showSupportLinks).
// The footer style is a plain style tag: style-src already allows 'unsafe-inline' (a nonce there would turn it off).
function buildHtml(uris, cspSource, nonce = crypto.randomBytes(16).toString('base64'), opts) {
  const footer = footerHtml('pptxViewer', !(opts && opts.showSupportLinks === false), 'zovo-support-fixed');
  const csp = buildCsp(cspSource, nonce);
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="${escapeAttr(csp)}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="${escapeAttr(uris.css)}">
${footer ? `<style>${FOOTER_CSS}</style>` : ''}
<title>PPTX Slide Viewer</title>
</head>
<body>
<div id="toolbar" role="toolbar">
  <button id="sidebarBtn" title="Toggle sidebar" aria-label="Toggle sidebar">&#9776;</button>
  <span id="fileName"></span>
  <span class="sep"></span>
  <button id="prevBtn" title="Previous slide (Left arrow)" aria-label="Previous slide">&#8249;</button>
  <input id="slideInput" type="text" inputmode="numeric" aria-label="Slide number" value="1"><span id="slideCount">/ 0</span>
  <button id="nextBtn" title="Next slide (Right arrow)" aria-label="Next slide">&#8250;</button>
  <span class="sep"></span>
  <span id="zoomGroup" class="group">
    <button id="zoomOutBtn" title="Zoom out (-)" aria-label="Zoom out">&#8722;</button>
    <select id="zoomSelect" title="Zoom (Ctrl or Cmd + mouse wheel also zooms)" aria-label="Zoom">
      <option value="fit">Fit slide</option>
      <option value="width">Fit width</option>
      <option value="0.5">50%</option>
      <option value="0.75">75%</option>
      <option value="1">100%</option>
      <option value="1.25">125%</option>
      <option value="1.5">150%</option>
      <option value="2">200%</option>
      <option value="3">300%</option>
      <option value="4">400%</option>
      <option value="custom" hidden></option>
    </select>
    <button id="zoomInBtn" title="Zoom in (+)" aria-label="Zoom in">+</button>
  </span>
  <button id="scrollBtn" title="Scroll view: all slides in one column" aria-pressed="false">Scroll</button>
  <span class="sep"></span>
  <input id="searchInput" type="search" placeholder="Search slides and notes" aria-label="Search slides and notes">
  <span id="searchCount"></span>
  <span class="grow"></span>
  <button id="notesBtn" title="Show or hide speaker notes">Notes</button>
  <button id="mdBtn" title="Copy this slide's text and notes as Markdown">Copy as Markdown</button>
  <button id="mdAllBtn" title="Copy all slides as Markdown">Copy all</button>
</div>
<div id="main">
  <aside id="sidebar">
    <div id="tabs" role="tablist">
      <button id="tabSlides" role="tab" class="active">Slides</button>
      <button id="tabOutline" role="tab">Outline</button>
    </div>
    <div id="thumbs" role="list"></div>
    <div id="outline" hidden></div>
  </aside>
  <section id="stageWrap">
    <div id="status" role="status"></div>
    <div id="notice" hidden></div>
    <div id="stage"></div>
    <div id="notes" aria-label="Speaker notes"></div>
  </section>
</div>
<div id="renderHost" aria-hidden="true"></div>
${footer}
<script nonce="${nonce}" src="${escapeAttr(uris.preview)}"></script>
<script nonce="${nonce}" src="${escapeAttr(uris.core)}"></script>
<script nonce="${nonce}" src="${escapeAttr(uris.viewer)}"></script>
</body>
</html>`;
}

module.exports = { buildHtml, buildCsp, escapeAttr };
