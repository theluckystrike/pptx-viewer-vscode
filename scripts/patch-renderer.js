'use strict';
// Fixes for pptx-preview 1.0.7 (pinned in package.json), applied to the UMD bundle at build time.
// Each patch must match exactly the expected number of times, or the build stops: a new
// pptx-preview version needs these checked again by hand.
//
// 1. Text style of placeholders. pptx-preview gave every placeholder the master's otherStyle and
//    only titles their titleStyle, so body text ("Click to add text" boxes, content placeholders)
//    lost the master's sizes, bullets and spacing (VER5 finding 1: 32 pt bullets drawn at 18 pt).
//    OOXML: title and ctrTitle use titleStyle; dt, ftr, sldNum and hdr use otherStyle; every other
//    placeholder (body, subTitle, obj, a placeholder with only an idx, ...) uses bodyStyle.
// 2. Vertical anchor. With no anchor anywhere in the chain, the text of a placeholder starts at the
//    top (the OOXML default). pptx-preview centred it, so body text sat too low on the slide.
//    Plain shapes keep pptx-preview's centring, which matches how PowerPoint saves them.

const RENDERER_VERSION = '1.0.7';

const STYLE_KEY = '(["ctrTitle","title"].includes(i)?"titleStyle":["dt","ftr","sldNum","hdr"].includes(i)?"otherStyle":"bodyStyle")';

const PATCHES = [
  {
    id: 'placeholder-text-style',
    count: 2, // paragraph props and run props
    find: /\(i\|\|o\)&&\(\["subTitle","ctrTitle","title"\]\.includes\(i\)&&Object\.assign\(n,Yn\.get\(r\.slideLayout\.slideMaster,\["textStyles","titleStyle","lvl"\.concat\(t\?\+t\+1:1\),"(props|defRPr)"\]\)\|\|\{\}\),/g,
    replace: (m, key) => `(i||o)&&(Object.assign(n,Yn.get(r.slideLayout.slideMaster,["textStyles",${STYLE_KEY},"lvl".concat(t?+t+1:1),"${key}"])||{}),`,
  },
  {
    id: 'placeholder-anchor-top',
    count: 1,
    find: /default:n\?u\.style\.top=e\.top\+"px":\(u\.style\.top=e\.top\+e\.h\/2\+"px",u\.style\.transform="translateY\(-50%\)"\)\}/g,
    replace: () => 'default:n||t.node&&(t.node.type||t.node.idx)?u.style.top=e.top+"px":(u.style.top=e.top+e.h/2+"px",u.style.transform="translateY(-50%)")}',
  },
  {
    // 3. Master lookup. Layouts and slides find their master placeholder by type, never by idx:
    //    a content placeholder with idx 2 got the master's date placeholder (also idx 2), so it was
    //    grey, small and centred. A placeholder with no type is a body; ctrTitle maps to title;
    //    pic, tbl, chart, obj, subTitle and the rest fall back to the master's body.
    id: 'master-placeholder-by-type',
    count: 1,
    find: /getNodeByType=function\(t\)\{return this\.nodes\.find\(\(function\(e\)\{return e\.type===t\}\)\)\},t\.prototype\.getNodeByIdx=function\(t\)\{return this\.nodes\.find\(\(function\(e\)\{return e\.idx===t\}\)\)\},t\.prototype\.getNodeInheritAttrsByType=function\(t,e\)\{\},/g,
    replace: () => 'getNodeByType=function(t){var n=this.nodes,r=n.find((function(e){return e.type===t}));if(r)return r;var k="ctrTitle"===t?"title":["dt","ftr","sldNum","hdr","title"].includes(t)?null:"body";return k?n.find((function(e){return e.type===k})):void 0},t.prototype.getNodeByIdx=function(t){return this.nodes.find((function(e){return e.type==="body"}))},t.prototype.getNodeInheritAttrsByType=function(t,e){},',
  },
  {
    // 4. Bullets. Paragraph and list-style parsers read buChar and buAutoNum but not buNone, so a
    //    "no bullet" lower in the chain could not cancel a bullet from the master. One kind of
    //    bullet also has to cancel the other kind.
    id: 'bullet-none-and-kind',
    count: 2,
    find: /Yn\.get\(t,\["a:buAutoNum","attrs","type"\]\)&&\(e\.buAutoNum=t\["a:buAutoNum"\]\.attrs\.type\),Yn\.get\(t,\["a:buChar","attrs","char"\]\)&&\(e\.buChar=t\["a:buChar"\]\.attrs\.char\),/g,
    replace: () => 'Yn.get(t,["a:buAutoNum","attrs","type"])&&(e.buAutoNum=t["a:buAutoNum"].attrs.type,e.buChar=""),Yn.get(t,["a:buChar","attrs","char"])&&(e.buChar=t["a:buChar"].attrs.char,e.buAutoNum=""),t&&"object"==typeof t&&"a:buNone"in t&&(e.buChar="",e.buAutoNum=""),',
  },
  {
    // 5. Master text styles (p:txStyles) were read without any bullet, so body text lost the
    //    master's bullets (VER5 deck: no bullet on any line).
    id: 'master-style-bullets',
    count: 1,
    find: /Yn\.get\(t,\["a:lnSpc","a:spcPct","attrs","val"\]\)&&\(e\.lineHeight=parseInt\(t\["a:lnSpc"\]\["a:spcPct"\]\.attrs\.val\)\/1e5\),e\}function Lr\(/g,
    replace: () => 'Yn.get(t,["a:lnSpc","a:spcPct","attrs","val"])&&(e.lineHeight=parseInt(t["a:lnSpc"]["a:spcPct"].attrs.val)/1e5),Yn.get(t,["a:buAutoNum","attrs","type"])&&(e.buAutoNum=t["a:buAutoNum"].attrs.type,e.buChar=""),Yn.get(t,["a:buChar","attrs","char"])&&(e.buChar=t["a:buChar"].attrs.char,e.buAutoNum=""),t&&"object"==typeof t&&"a:buNone"in t&&(e.buChar="",e.buAutoNum=""),e}function Lr(',
  },
  {
    // 6. A bullet character that is not one of the few Wingdings codes was drawn as a black square.
    //    Draw it as itself ("–", "o", "›", "✓" ...), and map a few more Wingdings and Symbol codes.
    id: 'bullet-char-fallback',
    count: 1,
    find: /textContent=\{n:"■",l:"●",u:"◆",p:"□","ü":"✔","Ø":"➢","•":"•"\}\[e\.buChar\]\|\|"■"/g,
    replace: () => 'textContent={n:"■",l:"●",u:"◆",p:"□","ü":"✔","Ø":"➢","•":"•","§":"▪",q:"❑",v:"❖","Ÿ":"•"}[e.buChar]||e.buChar',
  },
  {
    // 7. Fonts. Runs took their font only from a:ea (the East Asian font), so Latin text never got
    //    the deck's font, and theme fonts (+mn-lt, +mj-lt) were never looked up. Read a:latin first,
    //    resolve theme fonts, and keep the inherited font when a run names none.
    id: 'run-font-latin-and-theme',
    count: 1,
    find: /e\.typeface=Yn\.get\(t,\["a:ea","attrs","typeface"\]\),e\},t\}\(\);/g,
    replace: () => '(function(z){z&&(e.typeface=z)})(__zovoFace(t,this.node.theme)),e},t}();',
  },
  {
    id: 'table-font-latin-and-theme',
    count: 1,
    find: /return i&&\(e\.background=ir\(i,this\.theme,this\.ctx\)\),e\.typeface=Yn\.get\(t,\["a:ea","attrs","typeface"\]\),e\}/g,
    replace: () => 'return i&&(e.background=ir(i,this.theme,this.ctx)),(function(z){z&&(e.typeface=z)})(__zovoFace(t,this.theme)),e}',
  },
  {
    // Master and presentation default run styles (p:txStyles, defaultTextStyle) carry the theme font.
    id: 'style-font',
    count: 1,
    find: /var o=Yn\.get\(t,"a:solidFill"\);return o&&\(r\.color=ir\(o,e,n\)\),r\}function Dr\(/g,
    replace: () => 'var o=Yn.get(t,"a:solidFill"),zf=__zovoFace(t,e);return zf&&(r.typeface=zf),o&&(r.color=ir(o,e,n)),r}function Dr(',
  },
  {
    // 8. Every font gets fallbacks (Carlito for Calibri, Caladea for Cambria, ...) and a generic
    //    family, so a deck never falls back to the editor's UI font. The -4 % letter spacing that
    //    pptx-preview put on all Latin text without a font (meant for its Chinese default) is gone.
    id: 'font-stack',
    count: 1,
    find: /if\(a\.typeface\)switch\(l\.style\.fontFamily=a\.typeface,a\.typeface\)\{/g,
    replace: () => 'if(a.typeface)switch(l.style.fontFamily=__zovoStack(a.typeface),a.typeface){',
  },
  {
    id: 'no-latin-squeeze',
    count: 1,
    find: /else h\.test\(o\)&&\(l\.style\.letterSpacing=-\.04\*u\+"px"\)/g,
    replace: () => 'else l.style.fontFamily=__zovoStack("Calibri")',
  },
];

// Helpers the patches call. Plain functions, prepended to the bundle (the webview runs it as a classic script).
const HELPERS = `var __zovoFace=function(r,th){try{var g=function(o,k){for(var i=0;i<k.length;i++){if(o==null)return void 0;o=o[k[i]]}return o};var f=g(r,["a:latin","attrs","typeface"])||g(r,["a:ea","attrs","typeface"]);if(!f)return void 0;if(f.charAt(0)!=="+")return f;var s=th&&th.source;s=s&&(s["a:theme"]||s);var fs=g(s,["a:themeElements","a:fontScheme"]);var k=f.indexOf("+mj")===0?"a:majorFont":"a:minorFont";var x=/-ea$/.test(f)?"a:ea":/-cs$/.test(f)?"a:cs":"a:latin";return g(fs,[k,x,"attrs","typeface"])||g(fs,[k,"a:latin","attrs","typeface"])||void 0}catch(e){return void 0}};
var __zovoStack=function(f){var m={"calibri":'Carlito,"Segoe UI","Helvetica Neue",Arial',"calibri light":'Carlito,"Segoe UI Light","Helvetica Neue",Arial',"aptos":'Calibri,Carlito,"Segoe UI","Helvetica Neue",Arial',"aptos display":'Aptos,Calibri,Carlito,"Segoe UI","Helvetica Neue",Arial',"aptos narrow":'"Arial Narrow",Aptos,Calibri,Arial',"cambria":'Caladea,Georgia,"Times New Roman"',"arial":'"Liberation Sans",Helvetica',"times new roman":'"Liberation Serif",Times',"courier new":'"Liberation Mono",Courier',"consolas":'Menlo,"DejaVu Sans Mono","Courier New"',"segoe ui":'"Helvetica Neue",Arial',"century gothic":'"Avenir Next",Futura,Arial',"gill sans mt":'"Gill Sans",Arial',"trebuchet ms":'"Lucida Grande",Arial',"verdana":'"DejaVu Sans",Arial',"tahoma":'Verdana,Arial'};var n=String(f).replace(/["\\\;{}<>]/g,"").trim();var k=n.toLowerCase();var mono=/mono|courier|consolas|code/.test(k);var serif=!mono&&!/sans/.test(k)&&/serif|times|georgia|garamond|cambria|book|palatino|minion|baskerville|bodoni|caslon/.test(k);return(n?'"'+n+'",':"")+(m[k]?m[k]+",":"")+(mono?"monospace":serif?"serif":"sans-serif")};
`;

/** Returns the patched bundle and the list of applied patch ids. Throws when a patch does not match. */
function patchRenderer(src) {
  const applied = [];
  for (const p of PATCHES) {
    const n = (src.match(p.find) || []).length;
    if (n !== p.count) throw new Error(`pptx-preview patch "${p.id}" matched ${n} times, expected ${p.count}. Check the patch against the new renderer.`);
    src = src.replace(p.find, p.replace);
    applied.push(p.id);
  }
  return { src: HELPERS + src + `\n/* zovo patches: ${applied.join(', ')} */\n`, applied };
}

module.exports = { patchRenderer, PATCHES, RENDERER_VERSION, HELPERS };
