'use strict';
// Writes THIRD-PARTY-NOTICES.md from the license files in node_modules. Run: node scripts/notices.js
const fs = require('fs');
const path = require('path');
const root = path.resolve(__dirname, '..');
const pkgs = [
  ['fflate', 'bundled in out/extension.js (reads the deck ZIP)'],
  ['pptx-preview', 'media/vendor/pptx-preview.umd.js (draws slides); the file also bundles the packages below'],
  ['jszip', 'inside pptx-preview.umd.js (used under its MIT option)'],
  ['lodash', 'inside pptx-preview.umd.js'],
  ['uuid', 'inside pptx-preview.umd.js'],
  ['tslib', 'inside pptx-preview.umd.js'],
  ['echarts', 'inside pptx-preview.umd.js (draws charts)'],
  ['zrender', 'inside pptx-preview.umd.js (used by echarts)'],
];
let out = '# Third-party notices\n\nPPTX Slide Viewer is MIT licensed. It ships the following third-party code. Versions are the ones npm installed next to pptx-preview; its prebuilt file may embed slightly different versions of the same packages.\n';
for (const [name, where] of pkgs) {
  const dir = path.join(root, 'node_modules', name);
  const pj = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8'));
  out += `\n## ${name} ${pj.version}\n\nLicense: ${pj.license}. Where: ${where}.\n`;
  if (pj.homepage || pj.repository) out += `Source: ${pj.homepage || (pj.repository.url || pj.repository)}\n`;
  for (const f of fs.readdirSync(dir)) {
    if (/^(licen[cs]e|notice)(\.(md|markdown|txt))?$/i.test(f)) out += `\n### ${f}\n\n\`\`\`text\n${fs.readFileSync(path.join(dir, f), 'utf8').trim()}\n\`\`\`\n`;
  }
  if (name === 'pptx-preview') out += '\nThe npm package states license ISC. Its README says the published npm package is free to use, including commercial use. Its source code is not public, so this extension ships the published build unchanged.\n';
}
fs.writeFileSync(path.join(root, 'THIRD-PARTY-NOTICES.md'), out);
console.log('wrote THIRD-PARTY-NOTICES.md', out.length);
