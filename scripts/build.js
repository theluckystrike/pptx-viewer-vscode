'use strict';
// Bundle the extension host code and copy the webview renderer. Run: node scripts/build.js
const path = require('path');
const fs = require('fs');
const root = path.resolve(__dirname, '..');

require('esbuild').buildSync({
  entryPoints: [path.join(root, 'src/extension.js')],
  bundle: true,
  platform: 'node',
  target: 'node18',
  format: 'cjs',
  external: ['vscode'],
  outfile: path.join(root, 'out/extension.js'),
  minify: true,
  legalComments: 'eof',
});

const vendor = path.join(root, 'media/vendor');
fs.mkdirSync(vendor, { recursive: true });
const { patchRenderer, RENDERER_VERSION } = require('./patch-renderer');
const version = require(path.join(root, 'node_modules/pptx-preview/package.json')).version;
if (version !== RENDERER_VERSION) throw new Error(`pptx-preview is ${version}; scripts/patch-renderer.js was written for ${RENDERER_VERSION}.`);
const { src, applied } = patchRenderer(fs.readFileSync(path.join(root, 'node_modules/pptx-preview/dist/pptx-preview.umd.js'), 'utf8'));
fs.writeFileSync(path.join(vendor, 'pptx-preview.umd.js'), src);
fs.writeFileSync(path.join(vendor, 'VERSION'), version + ' + ' + applied.join(', ') + '\n');
console.log('built out/extension.js and media/vendor/pptx-preview.umd.js (' + applied.length + ' renderer patches)');
