// Turning a file into something to read and pace: every document becomes a list of units (pages or slides),
// each with a word count so the pacer can give denser pages more time.
//
//  pdf    → its pages, drawn by pdf.js
//  pptx   → its slides, drawn from the file's own layout (pptx.js)
//  images → one per page (a folder of exported slides, scans, photos of notes)
//  docx, html, epub, md, txt → reflowed in the reading theme and cut into pages of about a printed page each,
//          at paragraph boundaries, starting a new page at big headings
import { parsePptx } from './pptx.js';
import { store, uid } from './store.js';
import { readLocal, fileUrl } from './native.js';

export const ACCEPT = '.pdf,.pptx,.docx,.html,.htm,.xhtml,.epub,.md,.markdown,.txt,.text,.png,.jpg,.jpeg,.webp,.gif,.svg,.avif';
const IMG = /^(png|jpe?g|webp|gif|svg|avif)$/;
export const extOf = (name) => (name.match(/\.([a-z0-9]+)$/i)?.[1] || '').toLowerCase();
export function kindOf(name) {
  const e = extOf(name);
  if (e === 'pdf') return 'pdf';
  if (e === 'pptx') return 'slides';
  if (IMG.test(e)) return 'images';
  if (/^(docx|html?|xhtml|epub|md|markdown|txt|text)$/.test(e)) return 'flow';
  if (e === 'ppt' || e === 'doc' || e === 'key' || e === 'pages') return 'old';
  return null;
}
export const LABEL = { pdf: 'PDF', pptx: 'Slides', docx: 'Word', html: 'Web', htm: 'Web', xhtml: 'Web', epub: 'Book', md: 'Notes', markdown: 'Notes', txt: 'Text', text: 'Text' };
export const labelOf = (d) => (d.kind === 'images' ? 'Images' : LABEL[d.ext] || d.ext.toUpperCase());

const countWords = (s) => (s.match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu) || []).length;
export const baseName = (n) => n.replace(/\.[a-z0-9]+$/i, '').replace(/[_]+/g, ' ').trim();

// ---------- loaders for the big libraries, only when needed ----------
const loaded = {};
function script(src, global) {
  if (window[global]) return Promise.resolve(window[global]);
  loaded[src] ??= new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = src; s.onload = () => res(window[global]); s.onerror = () => { delete loaded[src]; rej(new Error(`Couldn't load ${src}`)); };
    document.head.append(s);
  });
  return loaded[src];
}
export const pdfjs = async () => {
  const lib = await script('vendor/pdf.min.js', 'pdfjsLib');
  lib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
  return lib;
};
const jszip = () => script('vendor/jszip.min.js', 'JSZip');
const mammoth = () => script('vendor/mammoth.browser.min.js', 'mammoth');

// ---------- one entry point ----------
// files: File[]; onStep(fileIndex, fraction, label)
// photos: { as: 'one' | 'each', name } — several photos become one collection (read in name order) or a document each
export async function importFiles(files, onStep, photos = {}) {
  const out = [];
  const errors = [];
  const images = [];
  const list = [...files];
  for (let i = 0; i < list.length; i++) {
    const f = list[i];
    const kind = kindOf(f.name);
    if (kind === 'images') { images.push(f); continue; }
    if (!kind) { errors.push(`${f.name}: this kind of file isn't supported`); continue; }
    if (kind === 'old') { errors.push(`${f.name}: save it as ${/ppt|key/.test(extOf(f.name)) ? '.pptx or PDF' : '.docx or PDF'} first`); continue; }
    try {
      onStep?.(i, 0, f.name);
      const d = await importOne(f, kind, (p) => onStep?.(i, p, f.name));
      out.push(d);
    } catch (e) {
      console.error(e);
      errors.push(`${f.name}: ${e.message || 'could not be read'}`);
    }
  }
  if (images.length) {
    try {
      if (photos.as === 'each' && images.length > 1) for (const im of images) out.push(await importImages([im]));
      else out.push(await importImages(images, photos.name));
    } catch (e) { errors.push(`Photos: ${e.message}`); }
  }
  return { docs: out, errors };
}

async function importOne(file, kind, onProgress) {
  const ext = extOf(file.name);
  const base = { id: uid(), name: baseName(file.name), file: file.name, ext, kind, size: file.size, added: Date.now(), pos: 0, read: [] };
  const buf = await file.arrayBuffer();
  const doc = { ...base, ...(await analyze(buf, kind, ext, base.id, onProgress)) };
  await store.putFile(base.id, new Blob([buf], { type: file.type || 'application/octet-stream' }));
  await store.putDoc(doc);
  return doc;
}

// What's worked out from a file: its units, words per unit, headings, shape and preview. Slides and reflowed
// documents also keep their parsed content (so they open instantly); the file itself is never needed for those.
async function analyze(buf, kind, ext, id, onProgress) {
  if (kind === 'pdf') return fromPdf(buf, onProgress);
  if (kind === 'slides') {
    await jszip();
    const deck = await parsePptx(buf, onProgress);
    if (!deck.slides.length) throw new Error('no slides found');
    await store.putContent(id, deck);
    return { units: deck.slides.length, unit: 'slide', words: deck.slides.map((s) => s.words), heads: deck.slides.map((s) => s.title), aspect: deck.w / deck.h };
  }
  const html = await flowHtml(buf, ext);
  const chunks = paginate(html.html);
  if (!chunks.length) throw new Error('nothing to read in it');
  await store.putContent(id, { chunks: chunks.map((c) => c.html) });
  const out = { units: chunks.length, unit: 'page', words: chunks.map((c) => c.words), heads: chunks.map((c) => c.head) };
  if (html.title && html.title.length < 120) out.title = html.title;
  return out;
}

// A document that lives in a folder on the computer (Windows app): read it from disk and work it out,
// keeping your progress and any name you gave it.
export async function indexLocal(doc, onProgress) {
  if (doc.kind === 'images') {
    const first = await readLocal(doc.paths[0]);
    const dims = await imageSize(first);
    const fields = { units: doc.paths.length, unit: 'page', words: doc.paths.map(() => 0), heads: doc.paths.map((p) => baseName(p.split(/[\\/]/).pop())), aspect: dims ? dims.w / dims.h : 0.77, thumb: await thumbFromBlob(first) };
    await store.putContent(doc.id, { paths: doc.paths });
    return { ...doc, ...fields, pending: false, error: null };
  }
  const buf = await (await readLocal(doc.path)).arrayBuffer();
  const fields = await analyze(buf, doc.kind, doc.ext, doc.id, onProgress);
  const next = { ...doc, ...fields, pending: false, error: null };
  // files in your folders keep their file names (that's how you organised them)
  // pages read that no longer exist (the file got shorter) are dropped
  next.read = (doc.read || []).filter((i) => i < next.units);
  next.pos = Math.min(doc.pos || 0, next.units - 1);
  return next;
}

async function importImages(files, name) {
  files.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  const id = uid();
  const blobs = files.map((f) => new Blob([f], { type: f.type }));
  await store.putContent(id, { images: blobs });
  const dims = await Promise.all(blobs.slice(0, 1).map(imageSize));
  const thumb = await thumbFromBlob(blobs[0]);
  const doc = {
    id, kind: 'images', collection: files.length > 1 || !!name, ext: extOf(files[0].name), name: name?.trim() || (files.length === 1 ? baseName(files[0].name) : `${baseName(files[0].name)} + ${files.length - 1} more`),
    file: files[0].name, size: files.reduce((s, f) => s + f.size, 0), added: Date.now(), pos: 0, read: [],
    units: files.length, unit: 'page', words: files.map(() => 0), heads: files.map((f) => baseName(f.name)), aspect: dims[0] ? dims[0].w / dims[0].h : 0.77, thumb,
  };
  await store.putDoc(doc);
  return doc;
}
// More photos into an existing collection: after what's there, or everything re-sorted by name.
export async function addToCollection(doc, files, { sort = false } = {}) {
  const c = await store.getContent(doc.id);
  const add = [...files].filter((f) => kindOf(f.name) === 'images').sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  if (!add.length) return { doc, added: 0 };
  let items = (c.images || []).map((b, i) => ({ b, h: doc.heads?.[i] || `Photo ${i + 1}` }));
  items.push(...add.map((f) => ({ b: new Blob([f], { type: f.type }), h: baseName(f.name) })));
  if (sort) items.sort((a, b) => a.h.localeCompare(b.h, undefined, { numeric: true }));
  await store.putContent(doc.id, { images: items.map((x) => x.b) });
  const next = { ...doc, collection: true, units: items.length, words: items.map(() => 0), heads: items.map((x) => x.h), size: (doc.size || 0) + add.reduce((s2, f) => s2 + f.size, 0) };
  if (!doc.thumb || sort) next.thumb = await thumbFromBlob(items[0].b);
  await store.putDoc(next);
  return { doc: next, added: add.length };
}
function imageSize(blob) {
  return new Promise((res) => { const u = URL.createObjectURL(blob); const im = new Image(); im.onload = () => { res({ w: im.naturalWidth, h: im.naturalHeight }); URL.revokeObjectURL(u); }; im.onerror = () => res(null); im.src = u; });
}
async function thumbFromBlob(blob) {
  const u = URL.createObjectURL(blob);
  try {
    const im = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = u; });
    const s = 360 / Math.max(im.naturalWidth, im.naturalHeight);
    const c = document.createElement('canvas'); c.width = Math.round(im.naturalWidth * s); c.height = Math.round(im.naturalHeight * s);
    c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
    return c.toDataURL('image/jpeg', 0.8);
  } catch { return null; } finally { URL.revokeObjectURL(u); }
}

// ---------- PDF ----------
async function fromPdf(buf, onProgress) {
  const lib = await pdfjs();
  const pdf = await lib.getDocument({ data: buf.slice(0), isEvalSupported: false }).promise;
  const n = pdf.numPages;
  const words = new Array(n).fill(0);
  const heads = new Array(n).fill('');
  let aspect = 0.77;
  let thumb = null;
  for (let i = 1; i <= n; i++) {
    const page = await pdf.getPage(i);
    if (i === 1) {
      const vp = page.getViewport({ scale: 1 });
      aspect = vp.width / vp.height;
      const s = 360 / Math.max(vp.width, vp.height);
      const v2 = page.getViewport({ scale: s * 1.5 });
      const c = document.createElement('canvas'); c.width = Math.ceil(v2.width); c.height = Math.ceil(v2.height);
      const ctx = c.getContext('2d', { alpha: false }); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: v2 }).promise;
      thumb = c.toDataURL('image/jpeg', 0.82);
    }
    try {
      const tc = await page.getTextContent();
      let text = '';
      let big = { size: 0, s: '' };
      for (const it of tc.items) {
        text += it.str + ' ';
        const h = Math.abs(it.transform?.[3] || 0);
        if (it.str.trim().length > 2 && h > big.size + 0.5) big = { size: h, s: it.str.trim() };
      }
      words[i - 1] = countWords(text);
      heads[i - 1] = big.s.slice(0, 90);
    } catch {}
    page.cleanup();
    if (i % 4 === 0 || i === n) { onProgress?.(i / n); await new Promise((r) => setTimeout(r)); }
  }
  // the document's outline, when it has one, names pages better than the largest text does
  try {
    const outline = await pdf.getOutline();
    if (outline?.length) {
      const named = new Array(n).fill('');
      const walk = async (items, depth) => {
        for (const it of items) {
          try {
            const dest = typeof it.dest === 'string' ? await pdf.getDestination(it.dest) : it.dest;
            if (dest?.[0]) { const idx = await pdf.getPageIndex(dest[0]); if (idx >= 0 && idx < n && !named[idx]) named[idx] = it.title; }
          } catch {}
          if (depth < 2 && it.items?.length) await walk(it.items, depth + 1);
        }
      };
      await walk(outline, 0);
      let last = '';
      for (let i = 0; i < n; i++) { if (named[i]) last = named[i]; if (last) heads[i] = last; }
    }
  } catch {}
  const info = await pdf.getMetadata().catch(() => null);
  pdf.destroy();
  const title = info?.info?.Title?.trim();
  return { units: n, unit: 'page', words, heads, aspect, thumb, ...(title && title.length > 3 && title.length < 120 && !/^(untitled|microsoft|slide)/i.test(title) ? { title } : {}) };
}

// ---------- reflowed documents ----------
async function flowHtml(buf, ext) {
  if (ext === 'docx') {
    const m = await mammoth();
    const r = await m.convertToHtml({ arrayBuffer: buf }, { styleMap: ['p[style-name=\'Title\'] => h1.title:fresh', 'p[style-name=\'Subtitle\'] => p.subtitle:fresh'] });
    const title = (r.value.match(/<h1[^>]*>(.*?)<\/h1>/i)?.[1] || '').replace(/<[^>]+>/g, '');
    return { html: r.value, title };
  }
  const text = new TextDecoder('utf-8').decode(buf);
  if (ext === 'md' || ext === 'markdown') {
    const html = markdown(text);
    return { html, title: (text.match(/^#\s+(.+)$/m)?.[1] || '').trim() };
  }
  if (ext === 'txt' || ext === 'text') {
    const paras = text.replace(/\r\n?/g, '\n').split(/\n\s*\n/).map((p) => p.trim()).filter(Boolean);
    return { html: paras.map((p) => `<p>${escapeHtml(p).replace(/\n/g, '<br>')}</p>`).join(''), title: '' };
  }
  if (ext === 'epub') return epub(buf);
  return cleanHtml(text, null);
}

const escapeHtml = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

// A web page, without what isn't the reading: scripts, site navigation, ads, forms. What's left takes the
// reading theme, so colours and fonts from the page are dropped and its own layout is simplified.
const DROP = 'script,noscript,style,link,meta,iframe,object,embed,form,input,button,select,textarea,nav,aside,footer,dialog,template,canvas,video,audio,[role=navigation],[role=banner],[role=contentinfo],[aria-hidden=true],.ad,.ads,.advert,.cookie,.newsletter,.share,.social,.sidebar,#comments,.comments';
const KEEP_STYLE = /^(text-align|font-style|font-weight|text-decoration|vertical-align|width|max-width|margin-left|padding-left)$/;
export function cleanHtml(html, baseUrl) {
  const d = new DOMParser().parseFromString(html, 'text/html');
  const title = d.querySelector('meta[property="og:title"]')?.content || d.title || d.querySelector('h1')?.textContent || '';
  const base = baseUrl || d.querySelector('base')?.href || d.querySelector('link[rel=canonical]')?.href || d.querySelector('meta[property="og:url"]')?.content || null;
  let root = d.querySelector('article') || d.querySelector('main') || d.querySelector('[role=main]') || d.body;
  if (!root) return { html: '', title };
  // an <article> that's only a teaser: fall back to the whole body
  if (root !== d.body && countWords(root.textContent || '') < 80) root = d.body;
  root.querySelectorAll(DROP).forEach((e) => e.remove());
  if (root === d.body) root.querySelectorAll('header').forEach((h) => { if (countWords(h.textContent) < 40) h.remove(); });
  for (const el of root.querySelectorAll('*')) {
    for (const a of [...el.attributes]) {
      const n = a.name.toLowerCase();
      if (n.startsWith('on') || n === 'class' || n === 'id' || n === 'bgcolor' || n === 'color' || n === 'face' || n === 'background') el.removeAttribute(a.name);
      else if (n === 'style') {
        const keep = a.value.split(';').map((s) => s.trim()).filter((s) => KEEP_STYLE.test(s.split(':')[0]?.trim().toLowerCase()));
        keep.length ? el.setAttribute('style', keep.join(';')) : el.removeAttribute('style');
      } else if ((n === 'href' || n === 'src') && /^\s*javascript:/i.test(a.value)) el.removeAttribute(a.name);
    }
    if (base && (el.tagName === 'IMG' || el.tagName === 'A')) {
      const k = el.tagName === 'IMG' ? 'src' : 'href';
      const v = el.getAttribute(k);
      if (v && !/^(data:|blob:|#|[a-z]+:)/i.test(v)) { try { el.setAttribute(k, new URL(v, base).href); } catch {} }
      if (el.tagName === 'IMG') { el.removeAttribute('srcset'); el.setAttribute('loading', 'lazy'); }
    }
    if (el.tagName === 'A') { el.setAttribute('target', '_blank'); el.setAttribute('rel', 'noopener noreferrer'); }
    if (el.tagName === 'FONT') el.replaceWith(...el.childNodes);
  }
  return { html: root.innerHTML, title: title.trim() };
}

async function epub(buf) {
  const Z = await jszip();
  const zip = await Z.loadAsync(buf);
  const parse = (s, t = 'application/xml') => new DOMParser().parseFromString(s, t);
  const container = parse(await zip.file('META-INF/container.xml').async('string'));
  const opfPath = container.querySelector('rootfile')?.getAttribute('full-path');
  if (!opfPath) throw new Error('not a readable EPUB');
  const opf = parse(await zip.file(opfPath).async('string'));
  const dir = opfPath.includes('/') ? opfPath.replace(/[^/]+$/, '') : '';
  const manifest = {};
  for (const it of opf.getElementsByTagNameNS('*', 'item')) manifest[it.getAttribute('id')] = { href: it.getAttribute('href'), type: it.getAttribute('media-type') };
  const title = opf.getElementsByTagNameNS('*', 'title')[0]?.textContent || '';
  const resolve = (from, href) => {
    const parts = from.split('/'); parts.pop();
    for (const s of decodeURIComponent(href.split('#')[0]).split('/')) { if (s === '..') parts.pop(); else if (s !== '.') parts.push(s); }
    return parts.join('/');
  };
  let html = '';
  for (const ref of opf.getElementsByTagNameNS('*', 'itemref')) {
    const it = manifest[ref.getAttribute('idref')];
    if (!it || !/x?html/.test(it.type || it.href)) continue;
    const p = dir + it.href;
    const f = zip.file(p) || zip.file(decodeURIComponent(p));
    if (!f) continue;
    const doc = parse(await f.async('string'), 'text/html');
    for (const img of doc.querySelectorAll('img, image')) {
      const k = img.tagName.toLowerCase() === 'image' ? 'xlink:href' : 'src';
      const src = img.getAttribute(k) || img.getAttribute('href');
      if (!src) continue;
      const ip = resolve(p, src);
      const imf = zip.file(ip);
      if (!imf) continue;
      const ext = ip.split('.').pop().toLowerCase();
      const b64 = await imf.async('base64');
      const url = `data:image/${ext === 'jpg' ? 'jpeg' : ext === 'svg' ? 'svg+xml' : ext};base64,${b64}`;
      if (k === 'src') img.setAttribute('src', url);
      else { const n = doc.createElement('img'); n.src = url; img.closest('svg')?.replaceWith(n); }
    }
    const body = cleanHtml(`<body>${doc.body?.innerHTML || ''}</body>`, null).html;
    if (countWords(body.replace(/<[^>]+>/g, ' ')) || /<img/i.test(body)) html += `<section data-chapter>${body}</section>`;
  }
  return { html, title };
}

// ---------- cutting reflowed text into pages ----------
// About a printed page each (target words), only between blocks, with a fresh page at each chapter or big heading
// once the page has something on it. Lists and long tables are split between their items when they're too long.
export const PAGE_WORDS = 330;
export function paginate(html, target = PAGE_WORDS) {
  const host = document.createElement('div');
  host.innerHTML = html;
  // flatten wrappers that only hold other blocks (div > div > section …)
  const blocks = [];
  const BLOCK = /^(P|H[1-6]|UL|OL|DL|TABLE|PRE|BLOCKQUOTE|FIGURE|IMG|HR|DIV|SECTION|ARTICLE|HEADER)$/;
  const flatten = (el, chapter) => {
    for (const n of [...el.childNodes]) {
      if (n.nodeType === 3) { if (n.textContent.trim()) blocks.push({ el: Object.assign(document.createElement('p'), { textContent: n.textContent.trim() }), chapter }); continue; }
      if (n.nodeType !== 1) continue;
      const isWrap = /^(DIV|SECTION|ARTICLE|HEADER|MAIN|SPAN|CENTER)$/.test(n.tagName) && [...n.children].some((c) => BLOCK.test(c.tagName));
      if (isWrap) { flatten(n, chapter || n.hasAttribute('data-chapter')); chapter = false; continue; }
      blocks.push({ el: n, chapter });
      chapter = false;
    }
  };
  flatten(host, false);
  const wordsOf = (el) => countWords(el.textContent || '') + (el.tagName === 'IMG' || el.querySelector?.('img') ? 60 : 0) + (el.tagName === 'TABLE' ? 20 : 0);
  // break up long lists/tables so a page never holds far more than its share
  const pieces = [];
  for (const b of blocks) {
    const w = wordsOf(b.el);
    if (w > target * 1.4 && /^(UL|OL)$/.test(b.el.tagName) && b.el.children.length > 1) {
      let cur = b.el.cloneNode(false), cw = 0, start = Number(b.el.getAttribute('start') || 1), idx = 0, first = true;
      for (const li of [...b.el.children]) {
        const lw = wordsOf(li);
        if (cw && cw + lw > target) { pieces.push({ el: cur, w: cw, chapter: first && b.chapter }); first = false; cur = b.el.cloneNode(false); if (b.el.tagName === 'OL') cur.setAttribute('start', start + idx); cw = 0; }
        cur.append(li); cw += lw; idx++;
      }
      pieces.push({ el: cur, w: cw, chapter: first && b.chapter });
    } else if (w > target * 1.4 && b.el.tagName === 'TABLE' && b.el.querySelectorAll('tr').length > 4) {
      const rows = [...b.el.querySelectorAll('tr')];
      const head = b.el.querySelector('thead');
      let cur = null, cw = 0, first = true;
      const fresh = () => { const t = document.createElement('table'); if (head && !first) t.append(head.cloneNode(true)); t.append(document.createElement('tbody')); return t; };
      for (const tr of rows) {
        if (head?.contains(tr)) continue;
        const rw = wordsOf(tr);
        if (!cur) { cur = fresh(); if (head && first) cur.prepend(head.cloneNode(true)); }
        if (cw && cw + rw > target) { pieces.push({ el: cur, w: cw, chapter: first && b.chapter }); first = false; cur = fresh(); cw = 0; }
        cur.querySelector('tbody').append(tr); cw += rw;
      }
      if (cur) pieces.push({ el: cur, w: cw, chapter: first && b.chapter });
    } else pieces.push({ el: b.el, w, chapter: b.chapter });
  }
  const pages = [];
  let cur = [], cw = 0, head = '', lastHead = '';
  const flush = () => {
    if (!cur.length) return;
    const html = cur.map((e) => e.outerHTML).join('');
    pages.push({ html, words: cw, head: head || lastHead });
    if (head) lastHead = head;
    cur = []; cw = 0; head = '';
  };
  for (let i = 0; i < pieces.length; i++) {
    const { el, w, chapter } = pieces[i];
    const isHead = /^H[1-3]$/.test(el.tagName);
    const big = /^H[12]$/.test(el.tagName) || chapter;
    if (cur.length && big && cw >= target * 0.35) flush();
    else if (cur.length && cw + w > target * 1.15 && cw >= target * 0.5) flush();
    // a heading never ends a page on its own
    if (isHead && cur.length && cw > target * 0.8) flush();
    if (isHead && !head) head = el.textContent.trim().slice(0, 90);
    cur.push(el); cw += w;
  }
  flush();
  return pages;
}

// ---------- a small, careful Markdown reader ----------
export function markdown(src) {
  const lines = src.replace(/\r\n?/g, '\n').split('\n');
  const out = [];
  const inline = (s) => escapeHtml(s)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/!\[([^\]]*)\]\(([^)\s]+)[^)]*\)/g, '<img alt="$1" src="$2">')
    .replace(/\[([^\]]+)\]\(([^)\s]+)[^)]*\)/g, (m, t, u) => (/^\s*javascript:/i.test(u) ? t : `<a href="${u}" target="_blank" rel="noopener">${t}</a>`))
    .replace(/\*\*([^*]+)\*\*|__([^_]+)__/g, '<strong>$1$2</strong>')
    .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, '$1<em>$2</em>')
    .replace(/(^|\W)_([^_\s][^_]*)_(?=\W|$)/g, '$1<em>$2</em>')
    .replace(/~~([^~]+)~~/g, '<del>$1</del>');
  let i = 0;
  while (i < lines.length) {
    const l = lines[i];
    if (/^```/.test(l)) {
      const body = []; i++;
      while (i < lines.length && !/^```/.test(lines[i])) body.push(lines[i++]);
      i++; out.push(`<pre><code>${escapeHtml(body.join('\n'))}</code></pre>`); continue;
    }
    let m;
    if ((m = l.match(/^(#{1,6})\s+(.*?)\s*#*\s*$/))) { out.push(`<h${m[1].length}>${inline(m[2])}</h${m[1].length}>`); i++; continue; }
    if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(l)) { out.push('<hr>'); i++; continue; }
    if (/^>\s?/.test(l)) { const b = []; while (i < lines.length && /^>\s?/.test(lines[i])) b.push(lines[i++].replace(/^>\s?/, '')); out.push(`<blockquote>${markdown(b.join('\n'))}</blockquote>`); continue; }
    if (/^\s*\|.*\|\s*$/.test(l) && /^\s*\|?\s*:?-{2,}/.test(lines[i + 1] || '')) {
      const row = (s) => s.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
      const h = row(l); i += 2; const rows = [];
      while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) rows.push(row(lines[i++]));
      out.push(`<table><thead><tr>${h.map((c) => `<th>${inline(c)}</th>`).join('')}</tr></thead><tbody>${rows.map((r) => `<tr>${r.map((c) => `<td>${inline(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>`);
      continue;
    }
    if (/^\s*([-*+]|\d+[.)])\s+/.test(l)) {
      const ordered = /^\s*\d/.test(l);
      const items = [];
      while (i < lines.length && /^\s*([-*+]|\d+[.)])\s+/.test(lines[i])) {
        let t = lines[i++].replace(/^\s*([-*+]|\d+[.)])\s+/, '');
        while (i < lines.length && /^\s{2,}\S/.test(lines[i]) && !/^\s*([-*+]|\d+[.)])\s+/.test(lines[i])) t += ' ' + lines[i++].trim();
        const task = t.match(/^\[([ xX])\]\s+(.*)/);
        items.push(task ? `<li class="task"><input type="checkbox" disabled${task[1] !== ' ' ? ' checked' : ''}> ${inline(task[2])}</li>` : `<li>${inline(t)}</li>`);
      }
      out.push(ordered ? `<ol>${items.join('')}</ol>` : `<ul>${items.join('')}</ul>`);
      continue;
    }
    if (!l.trim()) { i++; continue; }
    const p = [];
    while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|```|>|\s*([-*+]|\d+[.)])\s+)/.test(lines[i])) p.push(lines[i++]);
    out.push(`<p>${inline(p.join(' '))}</p>`);
  }
  return out.join('\n');
}
