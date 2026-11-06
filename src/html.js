'use strict';
const crypto = require('crypto');

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

function buildHtml(uris, cspSource, nonce = crypto.randomBytes(16).toString('base64')) {
  const csp = buildCsp(cspSource, nonce);
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="${escapeAttr(csp)}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="${escapeAttr(uris.css)}">
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
<script nonce="${nonce}" src="${escapeAttr(uris.preview)}"></script>
<script nonce="${nonce}" src="${escapeAttr(uris.core)}"></script>
<script nonce="${nonce}" src="${escapeAttr(uris.viewer)}"></script>
</body>
</html>`;
}

module.exports = { buildHtml, buildCsp, escapeAttr };
