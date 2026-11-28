// PowerPoint (.pptx) without PowerPoint: the slides are read straight from the file's XML and laid out as
// real elements at their true positions, so text stays sharp at any size and takes the reading theme.
// What's read: slide size and order, theme colours, backgrounds (slide › layout › master), shapes with fills,
// outlines and rounded/ellipse geometry, text with sizes, weights, colours, alignment and bullets
// (inheriting from the layout, master and the master's text styles), pictures, groups and tables.

const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};
const EMU_PT = 12700;

const parseXml = (s) => new DOMParser().parseFromString(s, 'application/xml');
// children by local name (namespaces in these files are reliable, but prefixes vary between writers)
const kids = (el, name) => (el ? [...el.children].filter((c) => c.localName === name) : []);
const kid = (el, name) => (el ? [...el.children].find((c) => c.localName === name) || null : null);
const path = (el, ...names) => { let e = el; for (const n of names) { e = kid(e, n); if (!e) return null; } return e; };
const deep = (el, name) => (el ? el.getElementsByTagNameNS('*', name) : []);
const attr = (el, n) => (el ? el.getAttribute(n) : null);
const rid = (el, n = 'embed') => (el ? el.getAttributeNS(NS.r, n) || el.getAttribute('r:' + n) : null);

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/'); parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}
async function rels(zip, file) {
  const p = file.replace(/([^/]+)$/, '_rels/$1.rels');
  const f = zip.file(p);
  const map = {};
  if (!f) return map;
  const x = parseXml(await f.async('string'));
  for (const r of deep(x, 'Relationship')) map[attr(r, 'Id')] = { target: resolvePath(file, attr(r, 'Target')), type: attr(r, 'Type') || '', external: attr(r, 'TargetMode') === 'External' };
  return map;
}

// ---- colour ----
const clamp = (v, a = 0, b = 1) => Math.max(a, Math.min(b, v));
function hexToHsl(h) {
  const n = parseInt(h, 16); let r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b); let hh = 0, s = 0; const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min; s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    hh = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4; hh /= 6;
  }
  return [hh, s, l];
}
function hslToHex(h, s, l) {
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  let r, g, b;
  if (s === 0) r = g = b = l; else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
  return [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
}
const PRESET = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', darkGray: 'A9A9A9', lightGray: 'D3D3D3' };

function colourOf(el, ctx) {
  // el: an element that holds one colour child (srgbClr, schemeClr, sysClr, prstClr)
  if (!el) return null;
  const c = [...el.children].find((x) => /^(srgbClr|schemeClr|sysClr|prstClr|scrgbClr)$/.test(x.localName));
  if (!c) return null;
  let hex = null;
  if (c.localName === 'srgbClr') hex = attr(c, 'val');
  else if (c.localName === 'sysClr') hex = attr(c, 'lastClr') || (attr(c, 'val') === 'window' ? 'FFFFFF' : '000000');
  else if (c.localName === 'prstClr') hex = PRESET[attr(c, 'val')] || '000000';
  else if (c.localName === 'scrgbClr') hex = ['r', 'g', 'b'].map((k) => Math.round(clamp(Number(attr(c, k)) / 100000) * 255).toString(16).padStart(2, '0')).join('');
  else {
    let v = attr(c, 'val');
    if (v === 'phClr' && ctx.ph) return ctx.ph;
    v = ctx.map[v] || v;
    hex = ctx.scheme[v] || null;
  }
  if (!hex) return null;
  let alpha = 1;
  let [h, s, l] = hexToHsl(hex);
  for (const m of c.children) {
    const val = Number(attr(m, 'val')) / 100000;
    if (m.localName === 'lumMod') l *= val;
    else if (m.localName === 'lumOff') l += val;
    else if (m.localName === 'tint') l = l + (1 - l) * (1 - val);
    else if (m.localName === 'shade') l *= val;
    else if (m.localName === 'satMod') s *= val;
    else if (m.localName === 'alpha') alpha = val;
  }
  const out = '#' + hslToHex(h, clamp(s), clamp(l));
  return alpha < 1 ? `color-mix(in srgb, ${out} ${Math.round(alpha * 100)}%, transparent)` : out;
}

function fillOf(spPr, ctx) {
  if (!spPr) return undefined;
  if (kid(spPr, 'noFill')) return 'none';
  const sf = kid(spPr, 'solidFill');
  if (sf) return colourOf(sf, ctx);
  const gf = kid(spPr, 'gradFill');
  if (gf) {
    const stops = deep(gf, 'gs');
    const cols = [...stops].map((g) => `${colourOf(g, ctx) || '#fff'} ${Number(attr(g, 'pos')) / 1000}%`);
    const lin = kid(gf, 'lin');
    const ang = lin ? Number(attr(lin, 'ang')) / 60000 + 90 : 180;
    return cols.length ? `linear-gradient(${ang}deg, ${cols.join(', ')})` : undefined;
  }
  return undefined;
}

// ---- text styles ----
// A level's run defaults and paragraph defaults, from an <a:lvlNpPr>
function lvlStyle(lvl, ctx) {
  if (!lvl) return {};
  const o = {};
  if (attr(lvl, 'algn')) o.algn = attr(lvl, 'algn');
  if (attr(lvl, 'marL')) o.marL = Number(attr(lvl, 'marL')) / EMU_PT;
  if (attr(lvl, 'indent')) o.indent = Number(attr(lvl, 'indent')) / EMU_PT;
  const d = kid(lvl, 'defRPr');
  if (d) Object.assign(o, runStyle(d, ctx));
  if (kid(lvl, 'buNone')) o.bullet = null;
  const bc = kid(lvl, 'buChar');
  if (bc) o.bullet = attr(bc, 'char');
  if (kid(lvl, 'buAutoNum')) o.bullet = '#';
  const lnSpc = path(lvl, 'lnSpc', 'spcPct');
  if (lnSpc) o.lh = Number(attr(lnSpc, 'val')) / 100000;
  const spcB = path(lvl, 'spcBef', 'spcPts');
  if (spcB) o.spcB = Number(attr(spcB, 'val')) / 100;
  return o;
}
function runStyle(r, ctx) {
  const o = {};
  if (!r) return o;
  if (attr(r, 'sz')) o.sz = Number(attr(r, 'sz')) / 100;
  if (attr(r, 'b') != null) o.b = attr(r, 'b') === '1' || attr(r, 'b') === 'true';
  if (attr(r, 'i') != null) o.i = attr(r, 'i') === '1' || attr(r, 'i') === 'true';
  if (attr(r, 'u') && attr(r, 'u') !== 'none') o.u = true;
  if (attr(r, 'strike') && attr(r, 'strike') !== 'noStrike') o.s = true;
  if (attr(r, 'cap') === 'all') o.caps = true;
  const sf = kid(r, 'solidFill');
  if (sf) o.color = colourOf(sf, ctx);
  const latin = kid(r, 'latin');
  if (latin) {
    let tf = attr(latin, 'typeface');
    if (tf === '+mj-lt') tf = ctx.fonts.major; else if (tf === '+mn-lt') tf = ctx.fonts.minor;
    if (tf) o.font = tf;
  }
  return o;
}
function listStyles(lst, ctx) {
  const out = [];
  if (!lst) return out;
  for (let i = 1; i <= 9; i++) out[i - 1] = lvlStyle(kid(lst, `lvl${i}pPr`), ctx);
  return out;
}
const merge = (...ls) => {
  const out = [];
  for (let i = 0; i < 9; i++) out[i] = Object.assign({}, ...ls.map((l) => (l && l[i]) || {}));
  return out;
};

// ---- placeholders ----
function phOf(sp) {
  const ph = path(sp, 'nvSpPr', 'nvPr', 'ph') || path(sp, 'nvPicPr', 'nvPr', 'ph') || path(sp, 'nvGraphicFramePr', 'nvPr', 'ph');
  if (!ph) return null;
  return { type: attr(ph, 'type') || 'body', idx: attr(ph, 'idx') };
}
const phKind = (t) => (t === 'title' || t === 'ctrTitle' ? 'title' : t === 'subTitle' ? 'body' : t === 'body' || t === 'obj' || !t ? 'body' : 'other');
function findPh(tree, ph) {
  if (!tree || !ph) return null;
  const all = tree.phs;
  return all.find((p) => ph.idx != null && p.ph.idx === ph.idx && ph.type === p.ph.type)
    || all.find((p) => ph.idx != null && p.ph.idx === ph.idx)
    || all.find((p) => p.ph.type === ph.type)
    || all.find((p) => phKind(p.ph.type) === phKind(ph.type) && phKind(ph.type) !== 'other')
    || null;
}
function xfrmOf(sp) {
  const x = path(sp, 'spPr', 'xfrm') || path(sp, 'xfrm') || path(sp, 'grpSpPr', 'xfrm');
  if (!x) return null;
  const off = kid(x, 'off'), ext = kid(x, 'ext');
  if (!off || !ext) return null;
  return {
    x: Number(attr(off, 'x')) / EMU_PT, y: Number(attr(off, 'y')) / EMU_PT,
    w: Number(attr(ext, 'cx')) / EMU_PT, h: Number(attr(ext, 'cy')) / EMU_PT,
    rot: Number(attr(x, 'rot') || 0) / 60000, flipH: attr(x, 'flipH') === '1',
  };
}
// a parsed layout/master: its placeholders (with positions and text styles) and its decorations
async function readTree(zip, file, ctx) {
  const x = parseXml(await zip.file(file).async('string'));
  const r = await rels(zip, file);
  const spTree = path(x.documentElement, 'cSld', 'spTree');
  const phs = [];
  for (const sp of kids(spTree, 'sp')) {
    const ph = phOf(sp);
    if (ph) phs.push({ ph, xfrm: xfrmOf(sp), lst: listStyles(path(sp, 'txBody', 'lstStyle'), ctx), body: path(sp, 'txBody', 'bodyPr'), spPr: kid(sp, 'spPr') });
  }
  return { x, r, file, spTree, phs, bg: path(x.documentElement, 'cSld', 'bg') };
}

function bgOf(bg, ctx, rmap, media) {
  if (!bg) return null;
  const pr = kid(bg, 'bgPr');
  if (pr) {
    const bf = kid(pr, 'blipFill');
    if (bf) { const id = rid(kid(bf, 'blip')); const m = id && rmap[id]; if (m) return { image: media(m.target) }; }
    const f = fillOf(pr, ctx);
    if (f && f !== 'none') return { fill: f };
  }
  const ref = kid(bg, 'bgRef');
  if (ref) { const c = colourOf(ref, ctx); if (c) return { fill: c }; }
  return null;
}

// ---- reading the whole deck ----
export async function parsePptx(buf, onProgress) {
  const JSZip = window.JSZip;
  const zip = await JSZip.loadAsync(buf);
  const pres = parseXml(await zip.file('ppt/presentation.xml').async('string'));
  const presRels = await rels(zip, 'ppt/presentation.xml');
  const sz = deep(pres, 'sldSz')[0];
  const W = Number(attr(sz, 'cx') || 12192000) / EMU_PT, H = Number(attr(sz, 'cy') || 6858000) / EMU_PT;
  const order = [...deep(pres, 'sldId')].map((s) => presRels[rid(s, 'id')]?.target).filter(Boolean);

  // theme
  const scheme = {};
  const fonts = { major: 'Calibri Light', minor: 'Calibri' };
  const themeRel = Object.values(presRels).find((r) => r.type.endsWith('/theme'));
  const themeFile = themeRel?.target || 'ppt/theme/theme1.xml';
  if (zip.file(themeFile)) {
    const th = parseXml(await zip.file(themeFile).async('string'));
    const cs = deep(th, 'clrScheme')[0];
    if (cs) for (const c of cs.children) { const v = c.firstElementChild; if (v) scheme[c.localName] = attr(v, 'val') && v.localName === 'srgbClr' ? attr(v, 'val') : attr(v, 'lastClr') || attr(v, 'val'); }
    const mj = deep(th, 'majorFont')[0], mn = deep(th, 'minorFont')[0];
    if (mj) fonts.major = attr(kid(mj, 'latin'), 'typeface') || fonts.major;
    if (mn) fonts.minor = attr(kid(mn, 'latin'), 'typeface') || fonts.minor;
  }

  const media = {}; // zip path -> Blob
  const mediaRef = (p) => { if (!(p in media)) media[p] = null; return p; };
  const trees = new Map();
  const tree = async (file, ctx) => { if (!trees.has(file)) trees.set(file, readTree(zip, file, ctx)); return trees.get(file); };

  const slides = [];
  for (let i = 0; i < order.length; i++) {
    const file = order[i];
    if (!zip.file(file)) continue;
    const sx = parseXml(await zip.file(file).async('string'));
    const sr = await rels(zip, file);
    const layoutFile = Object.values(sr).find((r) => r.type.endsWith('/slideLayout'))?.target;
    const baseCtx = { scheme, fonts, map: { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' } };
    const layout = layoutFile && zip.file(layoutFile) ? await tree(layoutFile, baseCtx) : null;
    let master = null;
    if (layout) {
      const mf = Object.values(layout.r).find((r) => r.type.endsWith('/slideMaster'))?.target;
      if (mf && zip.file(mf)) master = await tree(mf, baseCtx);
    }
    // colour map from the master (bg1 → lt1 …)
    const cm = master && kid(master.x.documentElement, 'clrMap');
    const map = { ...baseCtx.map };
    if (cm) for (const a of cm.attributes) map[a.name] = a.value;
    const ctx = { scheme, fonts, map };
    // master text styles
    const txS = master && kid(master.x.documentElement, 'txStyles');
    const mStyles = {
      title: listStyles(kid(txS, 'titleStyle'), ctx),
      body: listStyles(kid(txS, 'bodyStyle'), ctx),
      other: listStyles(kid(txS, 'otherStyle'), ctx),
    };
    const bg = bgOf(path(sx.documentElement, 'cSld', 'bg'), ctx, sr, mediaRef)
      || (layout && bgOf(layout.bg, ctx, layout.r, mediaRef))
      || (master && bgOf(master.bg, ctx, master.r, mediaRef))
      || { fill: '#' + (scheme[map.bg1] || 'FFFFFF') };

    const shapes = [];
    const words = { n: 0 };
    let title = '';
    const texts = [];
    // decorations from the master and layout (logos, bars), unless the slide hides them
    const showMaster = attr(sx.documentElement, 'showMasterSp') !== '0' && (!layout || attr(layout.x.documentElement, 'showMasterSp') !== '0');
    const walk = (spTree, rmap, src, offset) => {
      for (const el of spTree.children) {
        const ph = phOf(el);
        if (src !== 'slide' && ph) continue; // placeholders on layouts/masters are only prompts
        if (el.localName === 'sp') {
          const s = shapeFrom(el, rmap, ctx, { layout, master, mStyles, ph, mediaRef, offset });
          if (s) {
            shapes.push(s);
            if (src === 'slide') {
              words.n += s.words || 0;
              if (ph && phKind(ph.type) === 'title' && !title) title = s.text || '';
              else if (s.text) texts.push(s.text);
            }
          }
        } else if (el.localName === 'pic') {
          const s = picFrom(el, rmap, mediaRef, offset, { layout, master, ph });
          if (s) shapes.push(s);
        } else if (el.localName === 'grpSp') {
          const gx = path(el, 'grpSpPr', 'xfrm');
          let off = offset;
          if (gx) {
            const o = kid(gx, 'off'), e = kid(gx, 'ext'), co = kid(gx, 'chOff'), ce = kid(gx, 'chExt');
            if (o && e && co && ce) {
              const sxx = Number(attr(e, 'cx')) / Math.max(1, Number(attr(ce, 'cx'))), syy = Number(attr(e, 'cy')) / Math.max(1, Number(attr(ce, 'cy')));
              const g = { ox: Number(attr(o, 'x')) / EMU_PT, oy: Number(attr(o, 'y')) / EMU_PT, cx: Number(attr(co, 'x')) / EMU_PT, cy: Number(attr(co, 'y')) / EMU_PT, sx: sxx, sy: syy };
              off = (b) => { const r = { ...b, x: g.ox + (b.x - g.cx) * g.sx, y: g.oy + (b.y - g.cy) * g.sy, w: b.w * g.sx, h: b.h * g.sy }; return offset ? offset(r) : r; };
            }
          }
          walk(el, rmap, src, off);
        } else if (el.localName === 'graphicFrame') {
          const s = tableFrom(el, ctx, offset, mStyles);
          if (s) { shapes.push(s); if (src === 'slide') { words.n += s.words || 0; texts.push(s.rows.map((r) => r.cells.map((c) => c.paras.map((p) => p.map((x) => x.t).join('')).join(' ')).join(' · ')).join('\n')); } }
        }
      }
    };
    if (showMaster && master) walk(master.spTree, master.r, 'master');
    if (layout) walk(layout.spTree, layout.r, 'layout');
    walk(path(sx.documentElement, 'cSld', 'spTree'), sr, 'slide');

    // speaker notes count toward nothing, but they're kept for the reader
    let notes = '';
    const nf = Object.values(sr).find((r) => r.type.endsWith('/notesSlide'))?.target;
    if (nf && zip.file(nf)) {
      const nx = parseXml(await zip.file(nf).async('string'));
      for (const sp of deep(nx, 'sp')) {
        const ph = phOf(sp);
        if (ph && ph.type === 'body') notes = [...deep(sp, 'p')].map((p) => [...deep(p, 't')].map((t) => t.textContent).join('')).filter(Boolean).join('\n');
      }
    }
    slides.push({ bg, shapes, words: words.n, title: title.trim(), notes, text: texts });
    onProgress?.((i + 1) / order.length);
  }
  // pull out the pictures that were used
  for (const p of Object.keys(media)) {
    const f = zip.file(p);
    if (!f) continue;
    const ext = p.split('.').pop().toLowerCase();
    const type = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf' }[ext] || 'application/octet-stream';
    media[p] = new Blob([await f.async('arraybuffer')], { type });
  }
  return { w: W, h: H, slides, media, fonts };
}

function shapeFrom(sp, rmap, ctx, { layout, master, mStyles, ph, offset }) {
  const lph = ph ? findPh(layout, ph) : null;
  const mph = ph ? findPh(master, lph?.ph || ph) : null;
  let box = xfrmOf(sp) || lph?.xfrm || mph?.xfrm;
  if (!box) return null;
  if (offset) box = offset(box);
  const spPr = kid(sp, 'spPr');
  // style references (shapes drawn with the theme's look)
  const style = kid(sp, 'style');
  let fill = fillOf(spPr, ctx);
  if (fill === undefined && style) { const fr = kid(style, 'fillRef'); if (fr && attr(fr, 'idx') !== '0') fill = colourOf(fr, ctx); }
  if (fill === undefined && ph) fill = fillOf(lph?.spPr, ctx) ?? fillOf(mph?.spPr, ctx);
  let line = null;
  const ln = kid(spPr, 'ln');
  if (ln && !kid(ln, 'noFill')) {
    const c = colourOf(kid(ln, 'solidFill'), ctx);
    if (c) line = { c, w: Math.max(0.5, Number(attr(ln, 'w') || 12700) / EMU_PT) };
  } else if (!ln && style) {
    const lr = kid(style, 'lnRef');
    if (lr && attr(lr, 'idx') !== '0') { const c = colourOf(lr, ctx); if (c) line = { c, w: 1 }; }
  }
  const geom = attr(path(spPr, 'prstGeom'), 'prst') || 'rect';
  const out = { t: 'sp', x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot, fill: fill && fill !== 'none' ? fill : null, line, geom };
  // text
  const tx = kid(sp, 'txBody');
  if (tx) {
    const kind = ph ? phKind(ph.type) : 'other';
    const styles = merge(kind === 'title' ? mStyles.title : kind === 'body' ? mStyles.body : mStyles.other, mph?.lst, lph?.lst, listStyles(kid(tx, 'lstStyle'), ctx));
    const bodyPr = kid(tx, 'bodyPr');
    const bp = (n) => attr(bodyPr, n) ?? attr(lph?.body, n) ?? attr(mph?.body, n);
    out.anchor = bp('anchor') || (kind === 'title' ? 'ctr' : 't');
    out.vert = bp('vert');
    const ins = (n, d) => Number(bp(n) ?? d) / EMU_PT;
    out.pad = [ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720), ins('lIns', 91440)];
    const fit = kid(bodyPr, 'normAutofit');
    out.scale = fit && attr(fit, 'fontScale') ? Number(attr(fit, 'fontScale')) / 100000 : 1;
    out.wrap = bp('wrap') !== 'none';
    const textStyle = style ? colourOf(kid(style, 'fontRef'), ctx) : null;
    const paras = [];
    let text = '';
    let words = 0;
    let num = 0;
    for (const p of kids(tx, 'p')) {
      const pPr = kid(p, 'pPr');
      const lvl = Number(attr(pPr, 'lvl') || 0);
      const own = lvlStyle(pPr, ctx);
      const st = { ...styles[lvl], ...own };
      // a drawn shape's own text colour (its style's fontRef) beats the master's general text style
      if (!ph && textStyle && !own.color) st.color = textStyle;
      const runs = [];
      for (const r of p.children) {
        if (r.localName === 'r' || r.localName === 'fld') {
          const t = kid(r, 't')?.textContent || '';
          if (!t) continue;
          runs.push({ t, ...runStyle(kid(r, 'rPr'), ctx) });
          text += t; words += (t.match(/\S+/g) || []).length;
        } else if (r.localName === 'br') runs.push({ br: true });
      }
      const end = runStyle(kid(p, 'endParaRPr'), ctx);
      const has = runs.some((r) => r.t);
      if (st.bullet === '#' && has) num++; else if (!has) num = 0;
      paras.push({ runs, lvl, algn: st.algn || 'l', marL: st.marL || 0, indent: st.indent || 0, bullet: has && kind !== 'title' ? (st.bullet === '#' ? `${num}.` : st.bullet) : null, sz: st.sz || end.sz || (kind === 'title' ? 40 : 18), b: st.b, i: st.i, color: st.color, font: st.font, lh: st.lh, spcB: st.spcB, caps: st.caps });
      text += '\n';
    }
    out.paras = paras;
    out.text = text.trim();
    out.words = words;
  }
  return out;
}

function picFrom(pic, rmap, mediaRef, offset, { layout, master, ph }) {
  let box = xfrmOf(pic) || (ph && (findPh(layout, ph)?.xfrm || findPh(master, ph)?.xfrm));
  if (!box) return null;
  if (offset) box = offset(box);
  const blip = path(pic, 'blipFill', 'blip');
  const id = rid(blip);
  const m = id && rmap[id];
  if (!m || m.external) return null;
  const src = path(pic, 'blipFill', 'srcRect');
  const crop = src ? ['l', 't', 'r', 'b'].map((k) => Number(attr(src, k) || 0) / 100000) : null;
  return { t: 'pic', x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot, flip: box.flipH, src: mediaRef(m.target), crop };
}

function tableFrom(gf, ctx, offset) {
  const tbl = deep(gf, 'tbl')[0];
  if (!tbl) return null;
  let box = xfrmOf(gf);
  if (!box) return null;
  if (offset) box = offset(box);
  const cols = [...deep(kid(tbl, 'tblGrid'), 'gridCol')].map((c) => Number(attr(c, 'w')) / EMU_PT);
  let words = 0;
  const rows = kids(tbl, 'tr').map((tr) => ({
    h: Number(attr(tr, 'h')) / EMU_PT,
    cells: kids(tr, 'tc').map((tc) => {
      const paras = [...deep(tc, 'p')].map((p) => [...deep(p, 'r')].map((r) => ({ t: kid(r, 't')?.textContent || '', ...runStyle(kid(r, 'rPr'), ctx) })));
      for (const p of paras) for (const r of p) words += (r.t.match(/\S+/g) || []).length;
      const tcPr = kid(tc, 'tcPr');
      return { paras, fill: fillOf(tcPr, ctx), span: Number(attr(tc, 'gridSpan') || 1), merged: attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1' };
    }),
  }));
  return { t: 'tbl', x: box.x, y: box.y, w: box.w, h: box.h, cols, rows, words };
}

// ---- drawing a slide ----
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fontStack = (f, deck) => `"${(f || deck.fonts.minor).replace(/"/g, '')}", "${deck.fonts.minor}", Calibri, Carlito, "Segoe UI", system-ui, sans-serif`;

function runCss(r, p, deck, scale) {
  const sz = (r.sz || p.sz) * scale;
  let s = `font-size:${sz.toFixed(2)}pt;`;
  if (r.b ?? p.b) s += 'font-weight:700;';
  if (r.i ?? p.i) s += 'font-style:italic;';
  if (r.u || r.s) s += `text-decoration:${[r.u && 'underline', r.s && 'line-through'].filter(Boolean).join(' ')};`;
  const c = r.color || p.color;
  if (c) s += `color:${c};`;
  if (r.font || p.font) s += `font-family:${fontStack(r.font || p.font, deck)};`;
  if (r.caps || p.caps) s += 'text-transform:uppercase;';
  return s;
}

export function renderSlide(deck, i, urls) {
  const s = deck.slides[i];
  const root = document.createElement('div');
  root.className = 'slide';
  root.style.width = `${deck.w}pt`;
  root.style.height = `${deck.h}pt`;
  if (s.bg?.image && urls[s.bg.image]) root.style.background = `url("${urls[s.bg.image]}") center / cover`;
  else if (s.bg?.fill) root.style.background = s.bg.fill;
  let html = '';
  for (const sh of s.shapes) {
    const pos = `left:${sh.x}pt;top:${sh.y}pt;width:${sh.w}pt;height:${sh.h}pt;${sh.rot ? `transform:rotate(${sh.rot}deg);` : ''}`;
    if (sh.t === 'pic') {
      const u = urls[sh.src];
      if (!u) continue;
      let img = `<img src="${u}" alt="" draggable="false" style="${sh.flip ? 'transform:scaleX(-1);' : ''}">`;
      if (sh.crop && sh.crop.some((v) => v)) {
        const [l, t, r, b] = sh.crop;
        const w = 100 / Math.max(0.01, 1 - l - r), h = 100 / Math.max(0.01, 1 - t - b);
        img = `<img src="${u}" alt="" draggable="false" style="position:absolute;width:${w}%;height:${h}%;left:${-l * w}%;top:${-t * h}%;max-width:none">`;
      }
      html += `<div class="sl-pic" style="${pos}">${img}</div>`;
    } else if (sh.t === 'tbl') {
      const tw = sh.cols.reduce((a, b) => a + b, 0) || 1;
      html += `<div class="sl-tbl" style="${pos}"><table>${sh.cols.length ? `<colgroup>${sh.cols.map((c) => `<col style="width:${(c / tw) * 100}%">`).join('')}</colgroup>` : ''}${sh.rows.map((r) => `<tr style="height:${r.h}pt">${r.cells.filter((c) => !c.merged).map((c) => `<td${c.span > 1 ? ` colspan="${c.span}"` : ''} style="${c.fill && c.fill !== 'none' ? `background:${c.fill};` : ''}">${c.paras.map((p) => `<p>${p.map((r) => `<span style="${runCss(r, { sz: 14 }, deck, 1)}">${esc(r.t)}</span>`).join('')}</p>`).join('')}</td>`).join('')}</tr>`).join('')}</table></div>`;
    } else {
      const radius = sh.geom === 'ellipse' ? 'border-radius:50%;' : sh.geom === 'roundRect' ? `border-radius:${Math.min(sh.w, sh.h) * 0.16}pt;` : '';
      const box = `${sh.fill ? `background:${sh.fill};` : ''}${sh.line ? `box-shadow:inset 0 0 0 ${sh.line.w}pt ${sh.line.c};` : ''}${radius}`;
      let inner = '';
      if (sh.paras?.length && sh.text) {
        const jc = sh.anchor === 'ctr' ? 'center' : sh.anchor === 'b' ? 'flex-end' : 'flex-start';
        const [pt, pr, pb, pl] = sh.pad;
        inner = `<div class="sl-tx" style="justify-content:${jc};padding:${pt}pt ${pr}pt ${pb}pt ${pl}pt;${sh.wrap ? '' : 'white-space:nowrap;'}${sh.vert && sh.vert !== 'horz' ? 'writing-mode:vertical-rl;' : ''}">${sh.paras.map((p) => {
          const al = { ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[p.algn] || 'left';
          const body = p.runs.map((r) => (r.br ? '<br>' : `<span style="${runCss(r, p, deck, sh.scale)}">${esc(r.t)}</span>`)).join('');
          const base = `font-size:${(p.sz * sh.scale).toFixed(2)}pt;${p.color ? `color:${p.color};` : ''}${p.font ? `font-family:${fontStack(p.font, deck)};` : ''}`;
          const bullet = p.bullet ? `<span class="sl-bu" style="${base}">${esc(p.bullet === '#' ? '•' : p.bullet)}</span>` : '';
          const ml = p.marL || (p.bullet ? 18 + p.lvl * 18 : p.lvl * 18);
          return `<p style="text-align:${al};${base}padding-left:${ml}pt;text-indent:${p.bullet ? 0 : p.indent}pt;${p.lh ? `line-height:${(p.lh * 1.2).toFixed(2)};` : ''}${p.spcB ? `margin-top:${p.spcB * sh.scale}pt;` : ''}">${bullet}${body || '&#8203;'}</p>`;
        }).join('')}</div>`;
      }
      if (!box && !inner) continue;
      html += `<div class="sl-sp" style="${pos}${box}">${inner}</div>`;
    }
  }
  root.innerHTML = html;
  return root;
}
