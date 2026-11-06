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
fs.copyFileSync(path.join(root, 'node_modules/pptx-preview/dist/pptx-preview.umd.js'), path.join(vendor, 'pptx-preview.umd.js'));
fs.writeFileSync(path.join(vendor, 'VERSION'), require(path.join(root, 'node_modules/pptx-preview/package.json')).version + '\n');
console.log('built out/extension.js and media/vendor/pptx-preview.umd.js');
