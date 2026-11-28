import { store } from './store.js';
import { importFiles, ACCEPT, labelOf } from './importers.js';
import { Pacer, weightsFor, suggestTime, clock, human, timeOfDay, METHODS, methodCfg, DEFAULTS, carryOver } from './pacer.js';
import { Viewer, loadDoc } from './viewer.js';
import { renderSlide } from './pptx.js';
import { isDesktop, pickFolder, scanFolder, reveal, pathId } from './native.js';
import { indexLocal, kindOf, extOf, baseName, addToCollection, buildText, textVersion } from './importers.js';
import { compile, MODES, WHERE } from './search.js';
import { THEME_DEFAULTS, applyReadingTheme, openThemeBar, themeBarOpen, blurPx, BLUR_DEF } from './themes.js';

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const reduced = () => matchMedia('(prefers-reduced-motion: reduce)').matches;
const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
const ICON = {
  check: '<svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
  dots: '<svg viewBox="0 0 24 24"><circle cx="5.5" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="18.5" cy="12" r="1.6"/></svg>',
  prev: '<svg viewBox="0 0 24 24"><path d="M15 5l-7 7 7 7"/></svg>',
  next: '<svg viewBox="0 0 24 24"><path d="M9 5l7 7-7 7"/></svg>',
  play: '<svg viewBox="0 0 24 24"><path d="M8 5.5v13l10.5-6.5Z"/></svg>',
  pause: '<svg viewBox="0 0 24 24"><rect x="6.5" y="5" width="3.8" height="14" rx="1.2"/><rect x="13.7" y="5" width="3.8" height="14" rx="1.2"/></svg>',
  close: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6 6 18"/></svg>',
  grip: '<svg viewBox="0 0 24 24"><path d="M8 7h.01M8 12h.01M8 17h.01M16 7h.01M16 12h.01M16 17h.01" stroke-width="3.2"/></svg>',
  folder: '<svg viewBox="0 0 24 24"><path d="M3.5 7.5A2 2 0 0 1 5.5 5.5h4l2 2h7a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2Z"/></svg>',
  folderPlus: '<svg viewBox="0 0 24 24"><path d="M3.5 7.5A2 2 0 0 1 5.5 5.5h4l2 2h7a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2Z"/><path d="M12 11v5M9.5 13.5h5"/></svg>',
  refresh: '<svg viewBox="0 0 24 24"><path d="M20 12a8 8 0 1 1-2.4-5.7L20 8.5"/><path d="M20 3.5v5h-5"/></svg>',
  chev: '<svg viewBox="0 0 24 24"><path d="M9 6l6 6-6 6"/></svg>',
  trash: '<svg viewBox="0 0 24 24"><path d="M4 7h16M9 7V4.5h6V7M6.5 7l1 13h9l1-13"/></svg>',
  pen: '<svg viewBox="0 0 24 24"><path d="M4 20h4L19 9l-4-4L4 16Z"/></svg>',
  book: '<svg viewBox="0 0 24 24"><path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H20v15H6.5A2.5 2.5 0 0 0 4 20.5Z"/><path d="M4 20.5A2.5 2.5 0 0 1 6.5 18H20v3H6.5"/></svg>',
  timer: '<svg viewBox="0 0 24 24"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l2.5 2.5M9.5 2.5h5"/></svg>',
  reset: '<svg viewBox="0 0 24 24"><path d="M4 12a8 8 0 1 0 2.4-5.7L4 8.5"/><path d="M4 3.5v5h5"/></svg>',
  chart: '<svg viewBox="0 0 24 24"><path d="M5 20V10M12 20V4M19 20v-7"/></svg>',
  coffee: '<svg viewBox="0 0 24 24"><path d="M4 9h13v5a5 5 0 0 1-5 5H9a5 5 0 0 1-5-5Z"/><path d="M17 10h1.5a2.5 2.5 0 0 1 0 5H17M8 3v3M12 3v3"/></svg>',
  plus: '<svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>',
  view: '<svg viewBox="0 0 24 24"><rect x="5" y="3.5" width="14" height="17" rx="2"/><path d="M9 8h6M9 12h6M9 16h3"/></svg>',
  zoom: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="6.5"/><path d="M16 16l4 4M8.5 11h5M11 8.5v5"/></svg>',
  gear: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M12 2.5v3M12 18.5v3M2.5 12h3M18.5 12h3M5.3 5.3l2.1 2.1M16.6 16.6l2.1 2.1M5.3 18.7l2.1-2.1M16.6 7.4l2.1-2.1"/></svg>',
  stop: '<svg viewBox="0 0 24 24"><rect x="6" y="6" width="12" height="12" rx="2.5"/></svg>',
  notes: '<svg viewBox="0 0 24 24"><path d="M5 4h14v12l-4 4H5Z"/><path d="M15 20v-4h4M8.5 9h7M8.5 13h4"/></svg>',
  minus: '<svg viewBox="0 0 24 24"><path d="M5 12h14"/></svg>',
  search: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="6.5"/><path d="M16 16l4 4"/></svg>',
  down: '<svg viewBox="0 0 24 24"><path d="M6 9l6 6 6-6"/></svg>',
  up: '<svg viewBox="0 0 24 24"><path d="M6 15l6-6 6 6"/></svg>',
};

// ================= preferences =================
const PKEY = 'study-pacer:prefs';
const PREF_DEFAULTS = {
  ...THEME_DEFAULTS, glassT: 50, blurL: BLUR_DEF, glassG: {},
  view: 'paged', fit: 'page', turn: 'slide', tapEdges: true, notes: true,
  fontSize: 18, flowW: 720, flowFont: 'serif', lineH: 1.65,
  autoPause: true, sound: true, haptics: true, wake: true, nudge: true, hudCompact: false,
  sort: 'recent', setup: null, customMethod: null,
  sMode: 'all', matchCase: false, wholeWord: false, near: 8, where: 'any', sScope: 'all',
};
let prefs;
try { prefs = { ...PREF_DEFAULTS, ...JSON.parse(localStorage.getItem(PKEY) || '{}') }; } catch { prefs = { ...PREF_DEFAULTS }; }
const savePrefs = () => { try { localStorage.setItem(PKEY, JSON.stringify(prefs)); } catch {} };

// ================= glass (ported from Findly) =================
const GLASS = { '--a-top': 76, '--a-bar': 60, '--a-menu': 82, '--a-sheet': 97, '--a-tbar': 94 };
const GLASS_GROUPS = {
  top: { vars: ['--a-top'], blur: '--blur-top' },
  bottom: { vars: ['--a-bar'], blur: '--blur-bottom' },
  menus: { vars: ['--a-menu'], blur: '--blur-pop' },
  sheets: { vars: ['--a-sheet', '--a-tbar'], blur: '--blur-sheet' },
};
// out of the box: a little see-through with the slightest blur; sheets a touch more solid
const GLASS_START = { top: { t: 62, b: 34 }, bottom: { t: 62, b: 38 }, menus: { t: 62, b: 38 }, sheets: { t: 58, b: 45 } };
const glassOf = (g) => { const o = prefs.glassG?.[g]; const s = GLASS_START[g]; return { t: o?.t ?? s?.t ?? prefs.glassT ?? 50, b: o?.b ?? s?.b ?? prefs.blurL ?? BLUR_DEF }; };
function applyGlass() {
  const root = document.documentElement.style;
  for (const [g, { vars, blur }] of Object.entries(GLASS_GROUPS)) {
    const { t: tt, b } = glassOf(g);
    const t = Math.max(0, Math.min(100, tt)) / 100;
    for (const k of vars) {
      const base = GLASS[k];
      const a = t <= 0.5 ? 100 + (base - 100) * (t / 0.5) : base + (6 - base) * ((t - 0.5) / 0.5);
      root.setProperty(k, `${a.toFixed(1)}%`);
    }
    root.setProperty(blur, `saturate(180%) blur(${blurPx(b).toFixed(1)}px)`);
  }
  const sa = parseFloat(root.getPropertyValue('--a-sheet')) || 100;
  root.setProperty('--a-sheet-card', `${(sa + (100 - sa) * 0.3).toFixed(1)}%`);
  root.setProperty('--bpx', `${blurPx(prefs.blurL ?? BLUR_DEF).toFixed(1)}px`);
}
function applyFlowPrefs() {
  const r = document.documentElement.style;
  r.setProperty('--flow-font', prefs.flowFont === 'sans' ? 'var(--ui)' : prefs.flowFont === 'mono' ? 'var(--mono)' : prefs.flowFont === 'plain' ? 'Verdana, "Atkinson Hyperlegible", Tahoma, sans-serif' : 'var(--read)');
  r.setProperty('--flow-lh', String(prefs.lineH || 1.65));
  document.body.classList.toggle('tap-edges', !!prefs.tapEdges);
}
const openThemes = () => { closeMenu(); openThemeBar(prefs, { save: savePrefs, glass: applyGlass, glassOf }); };
applyReadingTheme(prefs);
applyGlass();
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if ((prefs.rtheme || 'auto') === 'auto') applyReadingTheme(prefs); });

// ================= small UI pieces =================
function toast(html, { ms = 3200, actions = [], plain = false, id } = {}) {
  const host = $('#toasts');
  if (id) host.querySelector(`[data-id="${id}"]`)?.remove();
  const t = document.createElement('div');
  t.className = `toast${plain || !actions.length ? ' plain' : ''}`;
  if (id) t.dataset.id = id;
  t.innerHTML = `<span>${html}</span>${actions.map((a, i) => `<button class="btn${a.primary ? ' primary' : ''}" data-i="${i}">${esc(a.label)}</button>`).join('')}`;
  const close = () => { if (!t.isConnected || t.classList.contains('out')) return; t.classList.add('out'); setTimeout(() => t.remove(), 260); };
  t.onclick = (e) => { const b = e.target.closest('[data-i]'); if (b) { actions[Number(b.dataset.i)].run?.(); close(); } };
  host.append(t);
  if (ms) setTimeout(close, ms);
  t.close = close;
  return t;
}
let menuEl = null;
function closeMenu() {
  if (!menuEl) return;
  const m = menuEl; menuEl = null;
  m.classList.add('out'); setTimeout(() => m.remove(), 170);
}
function openMenu(anchor, items) {
  closeMenu();
  const m = document.createElement('div');
  m.className = 'menu glass-menu';
  m.innerHTML = items.map((it, i) => (it === '-' ? '<hr>' : `<button data-i="${i}" class="${it.danger ? 'danger' : ''}">${it.icon ? ICON[it.icon] || it.icon : ''}<span>${esc(it.label)}</span>${it.key ? `<span class="m-key">${esc(it.key)}</span>` : ''}</button>`)).join('');
  document.body.append(m);
  const r = anchor.getBoundingClientRect();
  const mw = m.offsetWidth, mh = m.offsetHeight;
  let left = Math.min(innerWidth - mw - 8, Math.max(8, r.right - mw));
  let top = r.bottom + 6;
  let oy = 'top';
  if (top + mh > innerHeight - 8) { top = Math.max(8, r.top - mh - 6); oy = 'bottom'; }
  m.style.left = `${left}px`; m.style.top = `${top}px`;
  m.style.setProperty('--ox', `${r.left + r.width / 2 - left}px ${oy}`);
  m.onclick = (e) => { const b = e.target.closest('[data-i]'); if (!b) return; const it = items[Number(b.dataset.i)]; closeMenu(); it.run?.(); };
  menuEl = m;
  setTimeout(() => document.addEventListener('pointerdown', function off(e) { if (menuEl === m && !m.contains(e.target)) closeMenu(); document.removeEventListener('pointerdown', off, true); }, true));
}

// Sheets: slide up from the bottom (phone) or float in the middle (wide screens); swipe down on the grab bar
// or from the top of the content to close. Nothing behind scrolls while one is open.
const sheets = [];
function openSheet({ title, body = '', foot = '', wide = false, onClose, cls = '' }) {
  const scrim = document.createElement('div');
  scrim.className = 'scrim';
  const sh = document.createElement('section');
  sh.className = `sheet${wide ? ' wide' : ''} ${cls}`;
  sh.setAttribute('role', 'dialog');
  sh.innerHTML = `<div class="grab"></div><header class="sh-head"><h3>${esc(title)}</h3><button class="icon-btn" data-close aria-label="Close">${ICON.close}</button></header><div class="sh-body">${body}</div>${foot ? `<footer class="sh-foot">${foot}</footer>` : ''}`;
  document.body.append(scrim, sh);
  document.body.classList.add('locked');
  const desk = () => innerWidth >= 760;
  const inF = desk() ? [{ opacity: 0, transform: 'translate(-50%, -48%) scale(0.95)' }, { opacity: 1, transform: 'translate(-50%, -50%)' }] : [{ transform: 'translate(-50%, 100%)' }, { transform: 'translate(-50%, 0)' }];
  sh.animate(inF, { duration: reduced() ? 1 : 460, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1.04)' });
  let closed = false;
  const close = (fast) => {
    if (closed) return; closed = true;
    const i = sheets.indexOf(api); if (i >= 0) sheets.splice(i, 1);
    if (!sheets.length) document.body.classList.remove('locked');
    scrim.classList.add('out');
    const cur = getComputedStyle(sh).transform;
    const outF = desk() ? [{ opacity: 1 }, { opacity: 0, transform: 'translate(-50%, -48%) scale(0.96)' }] : [{ transform: cur === 'none' ? 'translate(-50%, 0)' : cur }, { transform: 'translate(-50%, 105%)' }];
    sh.animate(outF, { duration: fast || reduced() ? 1 : 280, easing: 'cubic-bezier(0.4, 0, 1, 1)', fill: 'forwards' }).onfinish = () => { sh.remove(); scrim.remove(); };
    onClose?.();
  };
  scrim.onclick = () => close();
  sh.querySelector('[data-close]').onclick = () => close();
  // swipe to dismiss
  let drag = null;
  const bodyEl = sh.querySelector('.sh-body');
  sh.addEventListener('pointerdown', (e) => {
    if (desk() || e.pointerType === 'mouse') return;
    const onGrab = e.target.closest('.grab, .sh-head');
    if (!onGrab && (bodyEl.scrollTop > 0 || e.target.closest('input, textarea, .q-handle, [data-noswipe]'))) return;
    drag = { y: e.clientY, t: performance.now(), dy: 0, grab: !!onGrab };
  });
  sh.addEventListener('pointermove', (e) => {
    if (!drag) return;
    const dy = e.clientY - drag.y;
    if (!drag.on && dy > 10 && (drag.grab || bodyEl.scrollTop <= 0)) { drag.on = true; sh.setPointerCapture(e.pointerId); }
    if (drag.on) { drag.dy = Math.max(0, dy); sh.style.transform = `translate(-50%, ${drag.dy}px)`; }
  });
  const up = () => {
    if (!drag) return;
    const d = drag; drag = null;
    if (!d.on) return;
    const v = d.dy / (performance.now() - d.t);
    if (d.dy > 120 || v > 0.6) close();
    else { sh.animate([{ transform: `translate(-50%, ${d.dy}px)` }, { transform: 'translate(-50%, 0)' }], { duration: 320, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1.2)' }); sh.style.transform = ''; }
  };
  sh.addEventListener('pointerup', up); sh.addEventListener('pointercancel', up);
  const api = { el: sh, body: bodyEl, close, set foot(h) { const f = sh.querySelector('.sh-foot'); if (f) f.innerHTML = h; } };
  sheets.push(api);
  return api;
}
const sw = (k, on, extra = '') => `<label class="switch"><input type="checkbox" data-k="${k}" ${on ? 'checked' : ''} ${extra}><i></i></label>`;
const fillTrack = (r) => r.style.setProperty('--fillpct', `${((r.value - r.min) / (r.max - r.min)) * 100}%`);

// ================= sound and touch =================
let actx;
function chime(kind) {
  if (!prefs.sound) return;
  try {
    actx ??= new (window.AudioContext || window.webkitAudioContext)();
    if (actx.state === 'suspended') actx.resume();
    const seq = { due: [659, 784, 988], end: [784, 659], finish: [523, 659, 784, 1047], soft: [880] }[kind] || [660];
    const t0 = actx.currentTime + 0.02;
    seq.forEach((f, i) => {
      const o = actx.createOscillator(), g = actx.createGain();
      o.type = 'sine'; o.frequency.value = f;
      const t = t0 + i * 0.13;
      g.gain.setValueAtTime(0, t); g.gain.linearRampToValueAtTime(0.16, t + 0.02); g.gain.exponentialRampToValueAtTime(0.0001, t + 0.9);
      o.connect(g).connect(actx.destination); o.start(t); o.stop(t + 1);
    });
  } catch {}
}
const buzz = (p) => { if (prefs.haptics) try { navigator.vibrate?.(p); } catch {} };
let wake = null;
async function keepAwake(on) {
  try {
    if (on && prefs.wake && 'wakeLock' in navigator && document.visibilityState === 'visible') { if (!wake) { wake = await navigator.wakeLock.request('screen'); wake.addEventListener('release', () => { wake = null; }); } }
    else if (!on && wake) { await wake.release(); wake = null; }
  } catch {}
}

// ================= library =================
let docs = [];
const selected = new Set();
const docById = (id) => docs.find((d) => d.id === id);
const unitWord = (d, n = 2) => (d.kind === 'slides' ? (n === 1 ? 'slide' : 'slides') : d.kind === 'images' && (d.collection || d.local) ? (n === 1 ? 'photo' : 'photos') : n === 1 ? 'page' : 'pages');
const readCount = (d) => (d.read || []).length;

async function loadLibrary() {
  docs = await store.allDocs();
  renderLibrary();
}
function sortedDocs(list = docs) {
  const s = prefs.sort;
  list = [...list];
  const nat = (a, b) => a.name.localeCompare(b.name, undefined, { numeric: true });
  if (s === 'name') list.sort(nat);
  else if (s === 'progress') list.sort((a, b) => readCount(b) / (b.units || 1) - readCount(a) / (a.units || 1) || nat(a, b));
  else list.sort((a, b) => (b.opened || b.added) - (a.opened || a.added) || nat(a, b));
  return list;
}
// the order to read things in: folder by folder, then by name (Lesson 2 before Lesson 10)
function readingOrder(ids) {
  const key = (d) => (d.local ? `${folderLabel(d.dir[0])}/${d.rel}` : `\uffff${d.name}`);
  return ids.map(docById).filter(Boolean).sort((a, b) => key(a).localeCompare(key(b), undefined, { numeric: true })).map((d) => d.id);
}
function estimateFor(d) { return d.words?.length ? suggestTime(d.words.map((w) => ({ words: w, kind: d.kind })), speed) : 0; }

// ---- folders: where you are, and what's in it ----
let cwd = (() => { try { return JSON.parse(sessionStorage.getItem('sp:cwd') || '[]'); } catch { return []; } })();
const setCwd = (c) => { cwd = c; try { sessionStorage.setItem('sp:cwd', JSON.stringify(c)); } catch {} };
const sameDir = (a = [], b = []) => a.length === b.length && a.every((x, i) => x === b[i]);
const under = (dir = [], base) => base.every((x, i) => dir[i] === x);
// the first part of a folder document's dir is the attached folder's full path; show its own name
function folderLabel(seg) { const f = folders.find((x) => x.path === seg); return f ? f.name : seg; }
function folderTiles(base) {
  const map = new Map();
  for (const d of docs) {
    const dir = d.dir || [];
    if (dir.length <= base.length || !under(dir, base)) continue;
    const seg = dir[base.length];
    let t = map.get(seg);
    if (!t) { t = { seg, path: [...base, seg], docs: 0, sub: new Set(), thumbs: [], units: 0, read: 0 }; map.set(seg, t); }
    t.docs++;
    t.units += d.units || 0; t.read += readCount(d);
    if (dir.length > base.length + 1) t.sub.add(dir[base.length + 1]);
    if (d.thumb && t.thumbs.length < 3) t.thumbs.push(d.thumb);
  }
  // attached folders show at the top level even before anything in them is found
  if (!base.length) for (const f of folders) if (!map.has(f.path)) map.set(f.path, { seg: f.path, path: [f.path], docs: 0, sub: new Set(), thumbs: [], units: 0, read: 0, missing: f.missing });
  return [...map.values()].sort((a, b) => folderLabel(a.seg).localeCompare(folderLabel(b.seg), undefined, { numeric: true }));
}
const docsUnder = (base) => docs.filter((d) => under(d.dir || [], base) && (d.dir || []).length >= base.length);
let tiles = [];
function folderHtml(t, i) {
  const ids = docsUnder(t.path).map((d) => d.id);
  const all = ids.length && ids.every((id) => selected.has(id));
  const pct = t.units ? Math.round((t.read / t.units) * 100) : 0;
  const meta = t.missing ? 'Folder not found' : [t.docs ? plural(t.docs, 'document') : 'Empty', t.sub.size ? plural(t.sub.size, 'folder') : ''].filter(Boolean).join(' · ');
  return `<div class="card folder${all ? ' sel' : ''}" role="listitem" data-tile="${i}" style="--i:${Math.min(i, 16)}" tabindex="0">
    <div class="card-art folder-art">
      <span class="f-back"></span>
      <span class="f-pages">${(t.thumbs.length ? t.thumbs : [null, null]).slice(0, 3).map((th, k) => `<i style="--k:${k}">${th ? `<img src="${th}" alt="" draggable="false">` : ''}</i>`).join('')}</span>
      <span class="f-front">${ICON.folder}</span>
      ${pct ? `<span class="card-prog"><i style="width:${pct}%"></i></span>` : ''}
    </div>
    ${t.docs ? `<button class="card-check" data-check aria-label="Select everything in ${esc(folderLabel(t.seg))}">${ICON.check}</button>` : ''}
    <button class="card-menu" data-fmenu aria-label="More for ${esc(folderLabel(t.seg))}">${ICON.dots}</button>
    <div class="card-meta"><b>${esc(folderLabel(t.seg))}</b><small>${meta}${pct ? ` · ${pct}%` : ''}</small></div>
  </div>`;
}
function cardHtml(d, i) {
  if (d.pending || d.error) {
    return `<div class="card pending${selected.has(d.id) ? ' sel' : ''}" role="listitem" data-id="${d.id}" style="--i:${Math.min(i, 16)}" tabindex="0">
      <div class="card-art"><div class="pend-art${d.error ? ' err' : ''}"><b>${esc(labelOf(d))}</b></div><span class="card-badge ${d.kind}">${esc(labelOf(d))}</span></div>
      <button class="card-check" data-check aria-label="Select ${esc(d.name)}">${ICON.check}</button>
      <button class="card-menu" data-cmenu aria-label="More for ${esc(d.name)}">${ICON.dots}</button>
      <div class="card-meta"><b>${esc(d.name)}</b><small>${d.error ? `Couldn’t read: ${esc(d.error)}` : 'Reading…'}</small></div>
    </div>`;
  }
  const pct = d.units ? Math.round((readCount(d) / d.units) * 100) : 0;
  const wide = (d.aspect || 0.77) > 1.05;
  let art;
  if (d.thumb) art = `<img src="${d.thumb}" alt="" loading="lazy" draggable="false">`;
  else if (d.kind === 'slides') art = `<div class="mini-slide" data-mini="${d.id}"><div style="--ar:${d.aspect || 1.78}"></div></div>`;
  else art = `<div class="mini-flow" data-flow="${d.id}"><b>${esc(d.name)}</b></div>`;
  return `<div class="card${selected.has(d.id) ? ' sel' : ''}" role="listitem" data-id="${d.id}" style="--i:${Math.min(i, 16)}" tabindex="0">
    <div class="card-art${wide ? ' wide' : ''}${d.kind !== 'flow' ? ' tintable' : ''}">${art}
      <span class="card-badge ${d.kind}">${esc(labelOf(d))}</span>
      ${pct ? `<span class="card-prog"><i style="width:${pct}%"></i></span>` : ''}
    </div>
    <button class="card-check" data-check aria-label="Select ${esc(d.name)}">${ICON.check}</button>
    <button class="card-menu" data-cmenu aria-label="More for ${esc(d.name)}">${ICON.dots}</button>
    <div class="card-meta"><b>${esc(d.name)}</b><small>${d.units} ${unitWord(d, d.units)} · ~${human(estimateFor(d))}${pct ? ` · ${pct}%` : ''}</small></div>
  </div>`;
}
let libIO = null;
function renderLibrary() {
  const grid = $('#grid');
  // a folder that's gone (removed, or emptied) takes you back up
  while (cwd.length && !docs.some((d) => under(d.dir || [], cwd)) && !(cwd.length === 1 && folders.some((f) => f.path === cwd[0]))) cwd = cwd.slice(0, -1);
  setCwd(cwd);
  tiles = folderTiles(cwd);
  const here = sortedDocs(docs.filter((d) => sameDir(d.dir || [], cwd)));
  const anything = docs.length > 0 || folders.length > 0;
  $('#empty').hidden = anything;
  $('#libHead').hidden = !anything;
  grid.innerHTML = tiles.map(folderHtml).join('') + here.map((d, i) => cardHtml(d, i + tiles.length)).join('');
  if (!tiles.length && !here.length && anything) grid.innerHTML = '<p class="muted grid-empty">Nothing in this folder yet.</p>';
  grid.classList.toggle('selecting', selected.size > 0);
  $$('#sortSeg button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.sort === prefs.sort)));
  renderCrumbs();
  renderTree();
  // previews that need the document itself, drawn as they scroll into view
  libIO?.disconnect();
  libIO = new IntersectionObserver((ents) => {
    for (const e of ents) {
      if (!e.isIntersecting) continue;
      libIO.unobserve(e.target);
      e.target.dataset.mini ? miniSlide(e.target) : miniFlow(e.target);
    }
  }, { rootMargin: '300px' });
  $$('[data-mini], [data-flow]', grid).forEach((el) => libIO.observe(el));
  renderSelbar();
  renderHero();
}
function renderCrumbs() {
  const c = $('#crumbs');
  if (!cwd.length) { c.innerHTML = `<h2>${folders.length ? 'Library' : 'Your documents'}</h2>`; return; }
  const parts = [`<button data-crumb="0">Library</button>`, ...cwd.map((seg, i) => (i === cwd.length - 1 ? `<h2>${esc(folderLabel(seg))}</h2>` : `<button data-crumb="${i + 1}">${esc(folderLabel(seg))}</button>`))];
  c.innerHTML = parts.join(`<span class="crumb-sep">${ICON.chev}</span>`);
}
function goFolder(path, dir = 1) {
  setCwd(path);
  renderLibrary();
  $('.lib-main').scrollTop = 0;
  if (!reduced()) $('#grid').animate([{ opacity: 0, transform: `translateX(${dir * 18}px)` }, { opacity: 1, transform: 'none' }], { duration: 320, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1)' });
}
// the folder tree beside the library on wide windows
function renderTree() {
  const tree = $('#tree');
  const has = docs.some((d) => (d.dir || []).length) || folders.length;
  document.body.classList.toggle('has-tree', !!has);
  if (!has) { tree.innerHTML = ''; return; }
  prefs.treeOpen ||= {};
  const node = (base, depth) => folderTiles(base).map((t) => {
    const key = t.path.join('\u0001');
    const kids = t.sub.size > 0;
    const open = prefs.treeOpen[key] ?? depth === 0;
    const cur = sameDir(t.path, cwd), inside = !cur && under(cwd, t.path) && cwd.length > t.path.length;
    return `<div class="tn"><div class="tn-row${cur ? ' cur' : ''}${inside ? ' in' : ''}" style="--d:${depth}">
        ${kids ? `<button class="tn-tog${open ? ' open' : ''}" data-tog="${esc(key)}" aria-label="${open ? 'Collapse' : 'Expand'}">${ICON.chev}</button>` : '<span class="tn-tog"></span>'}
        <button class="tn-go" data-go="${esc(key)}">${ICON.folder}<span>${esc(folderLabel(t.seg))}</span><small>${t.docs || ''}</small></button>
      </div>${kids && open ? `<div class="tn-kids">${node(t.path, depth + 1)}</div>` : ''}</div>`;
  }).join('');
  const loose = docs.filter((d) => !(d.dir || []).length).length;
  tree.innerHTML = `<div class="tn-row${!cwd.length ? ' cur' : ''}" style="--d:0"><span class="tn-tog"></span><button class="tn-go" data-go="">${ICON.book}<span>Library</span><small>${loose || ''}</small></button></div>${node([], 0)}
    ${isDesktop ? `<button class="tn-add" data-act="addFolder">${ICON.folderPlus}<span>Add folder</span></button>` : ''}`;
}
$('#tree').addEventListener('click', (e) => {
  const tog = e.target.closest('[data-tog]');
  if (tog) { const k = tog.dataset.tog; prefs.treeOpen[k] = !tog.classList.contains('open'); savePrefs(); renderTree(); return; }
  const go = e.target.closest('[data-go]');
  if (go) { const path = go.dataset.go ? go.dataset.go.split('\u0001') : []; const dir = path.length >= cwd.length ? 1 : -1; if (path.length) prefs.treeOpen[go.dataset.go] = true; goFolder(path, dir); }
});
$('#crumbs').addEventListener('click', (e) => { const b = e.target.closest('[data-crumb]'); if (b) goFolder(cwd.slice(0, Number(b.dataset.crumb)), -1); });

// ---- folders on the computer (Windows app) ----
let folders = [];
const indexing = { queue: [], busy: false, total: 0, done: 0, first: null };
async function loadFolders() { folders = (await store.get('folders')) || []; }
async function saveFolders() { await store.set('folders', folders); }
async function addFolder() {
  const path = await pickFolder().catch((e) => { toast(esc(String(e))); return null; });
  if (!path) return;
  if (folders.some((f) => f.path.toLowerCase() === path.toLowerCase())) { toast('That folder is already in your library'); goFolder([folders.find((f) => f.path.toLowerCase() === path.toLowerCase()).path]); return; }
  // a folder inside one that's already attached is already covered
  const parent = folders.find((f) => path.toLowerCase().startsWith(f.path.toLowerCase().replace(/[\\/]$/, '') + '\\'));
  if (parent) { toast(`Already included in “${esc(parent.name)}”`); return; }
  let name = path.replace(/[\\/]+$/, '').split(/[\\/]/).pop() || path;
  if (folders.some((f) => f.name === name)) name = `${name} (${path.split(/[\\/]/).slice(-2, -1)[0] || folders.length + 1})`;
  folders.push({ path, name, added: Date.now() });
  await saveFolders();
  store.persist();
  await syncFolders({ only: path, announce: true });
  goFolder([path], 1);
}
async function removeFolder(path) {
  folders = folders.filter((f) => f.path !== path);
  await saveFolders();
  const gone = docs.filter((d) => d.root === path);
  for (const d of gone) { await store.deleteDoc(d.id); selected.delete(d.id); }
  docs = docs.filter((d) => d.root !== path);
  const ses = await store.get('session');
  if (ses && ses.units.some((u) => gone.some((d) => d.id === u.doc))) await store.del('session');
  if (cwd[0] === path) setCwd([]);
  renderLibrary();
  toast('Folder removed from the library. Nothing was deleted from your computer.', { ms: 4000 });
}
// Compare each folder with what the library has: add new files, re-read changed ones, drop deleted ones.
// Progress on a file is kept as long as it's still there (its id comes from its path).
let syncing = null;
function syncFolders(opts = {}) {
  if (!isDesktop || !folders.length) return Promise.resolve();
  if (syncing) return syncing;
  syncing = (async () => {
    let added = 0, changed = 0, removed = 0;
    for (const f of folders) {
      if (opts.only && f.path !== opts.only) continue;
      let entries;
      try { entries = await scanFolder(f.path); f.missing = false; } catch (e) { f.missing = true; continue; }
      const rootName = f.path;
      const want = new Map();
      const imgs = new Map();
      for (const e of entries) {
        const kind = kindOf(e.rel);
        if (!kind || kind === 'old') continue;
        const segs = e.rel.split('/');
        const file = segs.pop();
        const dir = [rootName, ...segs];
        if (kind === 'images') {
          const k = dir.join('/');
          let g = imgs.get(k);
          if (!g) { g = { dir, files: [] }; imgs.set(k, g); }
          g.files.push(e);
          continue;
        }
        const id = pathId(e.path);
        want.set(id, { id, local: true, root: f.path, path: e.path, rel: e.rel, dir, name: baseName(file), ext: extOf(file), kind, size: e.size, mtime: e.mtime });
      }
      // pictures in one folder read as one set, in name order (exported slides, scanned pages)
      for (const g of imgs.values()) {
        g.files.sort((a, b) => a.rel.localeCompare(b.rel, undefined, { numeric: true }));
        const id = pathId(g.dir.join('/') + '|images');
        const label = g.dir.length > 1 ? g.dir[g.dir.length - 1] : f.name;
        want.set(id, { id, local: true, root: f.path, paths: g.files.map((x) => x.path), rel: g.dir.slice(1).join('/') + '/', dir: g.dir, name: g.files.length === 1 ? baseName(g.files[0].rel.split('/').pop()) : label, ext: extOf(g.files[0].rel), kind: 'images', size: g.files.reduce((s, x) => s + x.size, 0), mtime: Math.max(...g.files.map((x) => x.mtime)) + g.files.length });
      }
      const have = new Map(docs.filter((d) => d.root === f.path).map((d) => [d.id, d]));
      for (const [id, w] of want) {
        const old = have.get(id);
        if (!old) {
          const d = { ...w, added: w.mtime || Date.now(), pos: 0, read: [], pending: true, units: 0, words: [], heads: [] };
          docs.push(d); await store.putDoc(d); queueIndex(d.id); added++;
        } else if (old.size !== w.size || old.mtime !== w.mtime || (w.paths && w.paths.join('|') !== (old.paths || []).join('|'))) {
          Object.assign(old, { ...w, name: old.customName ? old.name : w.name, pending: true, error: null });
          await store.putDoc(old); queueIndex(old.id); changed++;
        } else {
          if (old.rel !== w.rel || !sameDir(old.dir, w.dir)) { Object.assign(old, { rel: w.rel, dir: w.dir }); await store.putDoc(old); }
          if (old.pending) queueIndex(old.id);
        }
        have.delete(id);
      }
      for (const d of have.values()) { await store.deleteDoc(d.id); docs = docs.filter((x) => x !== d); selected.delete(d.id); removed++; }
    }
    await saveFolders();
    renderLibrary();
    const bits = [added && `${added} new`, changed && `${changed} updated`, removed && `${removed} removed`].filter(Boolean);
    if (bits.length && (opts.announce || opts.quiet === false)) toast(`${ICON.folder} ${bits.join(' · ')}`, { ms: 2600 });
    else if (!bits.length && opts.announce === 'refresh') toast('Everything’s up to date', { ms: 1800 });
    syncing = null;
  })().catch((e) => { console.error(e); syncing = null; });
  return syncing;
}
function queueIndex(id, first = false) {
  if (indexing.queue.includes(id)) { if (first) { indexing.queue.splice(indexing.queue.indexOf(id), 1); indexing.queue.unshift(id); } }
  else { first ? indexing.queue.unshift(id) : indexing.queue.push(id); indexing.total++; }
  runIndex();
}
async function runIndex() {
  if (indexing.busy) return;
  indexing.busy = true;
  let lastPaint = 0;
  while (indexing.queue.length) {
    const id = indexing.queue.shift();
    const d = docById(id);
    renderStatus();
    if (d?.pending) {
      try {
        const next = await indexLocal(d);
        Object.assign(d, next);
        dropTexts([d.id]);
      } catch (e) { d.pending = false; d.error = e.message || 'unreadable'; }
      await store.putDoc(d);
      indexing.waiters?.get(id)?.forEach((r) => r());
      indexing.waiters?.delete(id);
    }
    indexing.done++;
    // refresh the cards now and then, not after every file
    if (performance.now() - lastPaint > 700 || !indexing.queue.length) { lastPaint = performance.now(); if ($('#reader').hidden) updateCards(); }
    await new Promise((r) => setTimeout(r, 0));
  }
  indexing.busy = false; indexing.total = 0; indexing.done = 0;
  renderStatus();
}
function updateCards() {
  // swap in finished cards without redrawing the whole grid (keeps your scroll and avoids flicker)
  for (const el of $$('#grid .card.pending')) {
    const d = docById(el.dataset.id);
    if (!d || d.pending) continue;
    const tmp = document.createElement('div');
    tmp.innerHTML = cardHtml(d, 0);
    const n = tmp.firstElementChild;
    n.style.animation = 'none';
    el.replaceWith(n);
    n.querySelector('.card-art').animate?.([{ opacity: 0.4 }, { opacity: 1 }], { duration: 300 });
    const m = n.querySelector('[data-mini], [data-flow]');
    if (m) libIO?.observe(m);
  }
  // folder tiles' counts and previews
  tiles = folderTiles(cwd);
  $$('#grid .card.folder').forEach((el) => { const t = tiles[Number(el.dataset.tile)]; if (!t) return; const tmp = document.createElement('div'); tmp.innerHTML = folderHtml(t, Number(el.dataset.tile)); const n = tmp.firstElementChild; n.style.animation = 'none'; el.replaceWith(n); });
  renderSelbar();
}
function renderStatus() {
  const st = $('#libStatus');
  const left = indexing.queue.length + (indexing.busy ? 1 : 0);
  if (!indexing.busy || !indexing.total) { st.hidden = true; return; }
  st.hidden = false;
  const done = Math.min(indexing.total, indexing.done);
  st.innerHTML = `<span class="spin"></span><span>Reading documents · ${done} of ${indexing.total}</span><i style="--p:${(done / indexing.total) * 100}%"></i>`;
  void left;
}
// make sure documents are worked out before reading or pacing them
async function ensureReady(ids) {
  const wait = ids.map(docById).filter((d) => d?.pending);
  if (!wait.length) return ids.filter((id) => !docById(id)?.error);
  const t = toast(`Getting ${wait.length === 1 ? esc(wait[0].name) : plural(wait.length, 'document')} ready…`, { ms: 0, plain: true, id: 'ready' });
  indexing.waiters ||= new Map();
  await Promise.all(wait.map((d) => new Promise((res) => { const l = indexing.waiters.get(d.id) || []; l.push(res); indexing.waiters.set(d.id, l); queueIndex(d.id, true); })));
  t.close();
  const bad = ids.map(docById).filter((d) => d?.error);
  if (bad.length) toast(`Couldn’t read ${esc(bad[0].name)}: ${esc(bad[0].error)}`, { ms: 5000 });
  return ids.filter((id) => docById(id) && !docById(id).error);
}
async function miniSlide(el) {
  const d = docById(el.dataset.mini);
  try {
    const x = await loadDoc(d);
    const box = el.firstElementChild;
    const s = renderSlide(x.deck, 0, x.urls);
    box.append(s);
    const fit = () => { const w = box.clientWidth; if (w) s.style.transform = `scale(${w / (x.deck.w * (96 / 72))})`; };
    fit(); new ResizeObserver(fit).observe(box);
  } catch {}
}
async function miniFlow(el) {
  const d = docById(el.dataset.flow);
  try {
    const c = await store.getContent(d.id);
    const tmp = document.createElement('div');
    tmp.innerHTML = c.chunks[0] || '';
    tmp.querySelectorAll('h1,h2,h3,img,table,pre').forEach((n) => n.remove());
    el.insertAdjacentHTML('beforeend', [...tmp.querySelectorAll('p, li')].slice(0, 8).map((p) => `<p>${esc(p.textContent.slice(0, 220))}</p>`).join(''));
  } catch {}
}
function renderSelbar() {
  const bar = $('#selbar');
  const was = !bar.hidden;
  const ids = [...selected].filter(docById);
  bar.hidden = !ids.length;
  if (!ids.length) return;
  const ds = ids.map(docById);
  const units = ds.reduce((a, d) => a + (d.units || 0), 0);
  const est = ds.reduce((a, d) => a + estimateFor(d), 0);
  $('#selCount').textContent = ids.length === 1 ? ds[0].name : `${ids.length} documents`;
  $('#selInfo').textContent = `${units} pages · about ${human(est)}`;
  if (!was) { bar.classList.remove('in'); void bar.offsetWidth; bar.classList.add('in'); }
}
async function renderHero() {
  const hero = $('#hero');
  const ses = await store.get('session');
  const hist = (await store.get('history')) || [];
  let html = '';
  if (ses && !ses.pacer?.s?.finished) {
    const p = Pacer.restore(ses);
    const sn = p.snapshot();
    const pct = sn.doneN / sn.n;
    const C = 2 * Math.PI * 27;
    const names = [...new Set(ses.units.map((u) => docById(u.doc)?.name).filter(Boolean))];
    html += `<div class="hero"><div class="hero-ring"><svg viewBox="0 0 66 66"><circle class="tr" cx="33" cy="33" r="27"/><circle class="ar" cx="33" cy="33" r="27" stroke-dasharray="${C}" stroke-dashoffset="${C * (1 - pct)}"/></svg><b>${Math.round(pct * 100)}%</b></div>
      <div class="hero-text"><small>Session in progress</small><h3>${esc(names.slice(0, 2).join(' + '))}${names.length > 2 ? ` + ${names.length - 2} more` : ''}</h3><p>${plural(sn.left, 'page')} to go · ${sn.open ? `${human(sn.work)} so far` : sn.remaining > 0 ? `${human(sn.remaining)} left` : 'out of time'}</p></div>
      <div class="hero-act"><button class="btn" data-act="endSaved">End</button><button class="btn primary" data-act="resume">${ICON.play}<span>Resume</span></button></div></div>`;
  }
  if (hist.length) {
    const wk = Date.now() - 7 * 864e5;
    const recent = hist.filter((h) => h.t > wk);
    const today = hist.filter((h) => new Date(h.t).toDateString() === new Date().toDateString());
    const sum = (a, k) => a.reduce((s, h) => s + (h[k] || 0), 0);
    const streak = (() => { let n = 0; const days = new Set(hist.map((h) => new Date(h.t).toDateString())); for (let d = new Date(); days.has(d.toDateString()); d.setDate(d.getDate() - 1)) n++; return n; })();
    html += `<div class="week"><div><b>${human(sum(today, 'work'))}</b><span>today</span></div><div><b>${human(sum(recent, 'work'))}</b><span>this week</span></div><div><b>${sum(recent, 'pages')}</b><span>pages this week</span></div>${speed.wpm ? `<div><b>${speed.wpm} wpm</b><span>your study pace</span></div>` : ''}${streak > 1 ? `<div><b>${streak} days</b><span>streak</span></div>` : ''}</div>`;
  }
  hero.innerHTML = html;
}

// importing
const filePick = $('#filePick');
filePick.accept = ACCEPT;
filePick.onchange = () => { if (filePick.files.length) doImport([...filePick.files]); filePick.value = ''; };
function askPhotos(n, first) {
  return new Promise((res) => {
    let done = false;
    const finish = (v) => { if (done) return; done = true; res(v); };
    const guess = baseName(first).replace(/[\s_-]*\(?\d+\)?$/, '').trim() || 'Photos';
    const s = openSheet({
      title: `${n} photos`,
      body: `<div class="group"><button class="row pick-row on" data-as="one"><span class="pick-dot"></span><div class="rt"><b>One collection</b><small>Read them in order like pages, and add more photos to it later</small></div></button>
        <div class="row" data-name-row><input class="field-in text" id="colName" value="${esc(guess)}" maxlength="120" placeholder="Collection name" aria-label="Collection name"></div>
        <button class="row pick-row" data-as="each"><span class="pick-dot"></span><div class="rt"><b>Separate photos</b><small>Each photo is its own document</small></div></button></div>
        ${collections().length ? `<div class="group-t">Or add them to a collection</div><div class="group">${collections().map((d) => `<button class="row pick-row" data-into="${d.id}"><span class="pick-dot"></span><div class="rt"><b>${esc(d.name)}</b><small>${plural(d.units, 'photo')}</small></div></button>`).join('')}</div>` : ''}`,
      foot: '<span class="grow"></span><button class="btn primary" data-ok>Import</button>',
      onClose: () => finish(null),
    });
    let choice = { as: 'one' };
    s.body.addEventListener('click', (e) => {
      const r = e.target.closest('.pick-row');
      if (!r) return;
      $$('.pick-row', s.body).forEach((x) => x.classList.toggle('on', x === r));
      choice = r.dataset.into ? { into: r.dataset.into } : { as: r.dataset.as };
      s.body.querySelector('[data-name-row]').hidden = choice.as !== 'one';
    });
    s.el.querySelector('[data-ok]').onclick = () => { if (choice.as === 'one') choice.name = $('#colName', s.el).value.trim() || guess; const v = choice; finish(v); s.close(); };
  });
}
const collections = () => docs.filter((d) => d.kind === 'images' && !d.local && d.collection);
async function addPhotos(d, files) {
  const t = toast(`Adding to ${esc(d.name)}…`, { ms: 0, plain: true, id: 'addp' });
  const { doc, added } = await addToCollection(d, files);
  t.close();
  Object.assign(d, doc);
  renderLibrary();
  toast(added ? `${ICON.check}Added ${plural(added, 'photo')} to ${esc(d.name)}` : 'No photos in what you picked', { ms: 2600 });
}
function pickPhotos(d) {
  const inp = document.createElement('input');
  inp.type = 'file'; inp.multiple = true; inp.accept = 'image/*,.png,.jpg,.jpeg,.webp,.gif,.avif';
  inp.onchange = () => { if (inp.files.length) addPhotos(d, [...inp.files]); };
  inp.click();
}
async function doImport(files) {
  store.persist();
  // several photos: one collection, separate photos, or into a collection you already have
  let photos = {};
  const imgs = files.filter((f) => kindOf(f.name) === 'images');
  if (imgs.length > 1 || (imgs.length === 1 && collections().length)) {
    photos = await askPhotos(imgs.length, imgs[0].name);
    if (!photos) return;
    if (photos.into) {
      await addPhotos(docById(photos.into), imgs);
      files = files.filter((f) => kindOf(f.name) !== 'images');
      if (!files.length) return;
      photos = {};
    }
  }
  const t = toast(`<span>Importing…</span> <span class="prog"><i></i></span>`, { ms: 0, plain: true, id: 'imp' });
  const bar = t.querySelector('.prog i');
  const lab = t.querySelector('span span') || t.firstElementChild.firstElementChild;
  const { docs: added, errors } = await importFiles(files, (i, p, name) => {
    bar.style.setProperty('--p', `${((i + p) / files.length) * 100}%`);
    lab.textContent = files.length > 1 ? `Reading ${i + 1} of ${files.length}` : `Reading ${name.length > 28 ? name.slice(0, 26) + '…' : name}`;
  }, photos);
  t.close();
  docs.push(...added);
  added.forEach((d) => selected.add(d.id));
  renderLibrary();
  if (added.length) toast(`${ICON.check}Added ${added.length === 1 ? esc(added[0].name) : plural(added.length, 'document')}`, { actions: [{ label: 'Pace it', primary: true, run: () => openSetup([...added.map((d) => d.id)]) }], ms: 5000 });
  for (const e of errors) toast(esc(e), { ms: 6000 });
}
// drag and drop, anywhere
let dragDepth = 0;
addEventListener('dragenter', (e) => { if (![...(e.dataTransfer?.types || [])].includes('Files')) return; dragDepth++; $('#dropVeil').hidden = false; });
addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; $('#dropVeil').hidden = true; } });
addEventListener('dragover', (e) => { if ([...(e.dataTransfer?.types || [])].includes('Files')) e.preventDefault(); });
addEventListener('drop', (e) => {
  if (!e.dataTransfer?.files?.length) return;
  e.preventDefault(); dragDepth = 0; $('#dropVeil').hidden = true;
  const onto = docById(e.target.closest?.('.card')?.dataset.id);
  if (onto && onto.kind === 'images' && !onto.local && reader.hidden) { addPhotos(onto, [...e.dataTransfer.files]); return; }
  if (!$('#reader').hidden) exitReader();
  doImport([...e.dataTransfer.files]);
});
// a share target / launch queue (installed app: "Open with Study Pacer")
if ('launchQueue' in window) window.launchQueue.setConsumer(async (p) => { if (p.files?.length) doImport(await Promise.all(p.files.map((h) => h.getFile()))); });

$('#grid').addEventListener('click', (e) => {
  const tile = e.target.closest('.card.folder');
  if (tile) {
    const t = tiles[Number(tile.dataset.tile)];
    if (!t) return;
    if (e.target.closest('[data-fmenu]')) { folderMenu(e.target.closest('[data-fmenu]'), t); return; }
    if (e.target.closest('[data-check]') || e.metaKey || e.ctrlKey || e.shiftKey) { toggleFolderSel(t, tile); return; }
    goFolder(t.path, 1);
    return;
  }
  const card = e.target.closest('.card');
  if (!card) return;
  const id = card.dataset.id;
  if (e.target.closest('[data-cmenu]')) { cardMenu(e.target.closest('[data-cmenu]'), id); return; }
  if (e.target.closest('[data-check]') || e.metaKey || e.ctrlKey || e.shiftKey || (selected.size && matchMedia('(hover: none)').matches)) { toggleSel(id, card); return; }
  openFree([id], card);
});
$('#grid').addEventListener('keydown', (e) => { const f = e.target.closest('.card.folder'); if (f && e.key === 'Enter') { const t = tiles[Number(f.dataset.tile)]; if (t) goFolder(t.path, 1); return; } const c = e.target.closest('.card:not(.folder)'); if (c && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); e.key === ' ' ? toggleSel(c.dataset.id, c) : openFree([c.dataset.id], c); } });
// long press selects (phones)
let lp;
$('#grid').addEventListener('pointerdown', (e) => { const c = e.target.closest('.card:not(.folder)'); if (!c || e.pointerType === 'mouse') return; lp = setTimeout(() => { buzz(12); toggleSel(c.dataset.id, c); c.addEventListener('click', (ev) => { ev.stopPropagation(); }, { capture: true, once: true }); }, 480); });
['pointerup', 'pointercancel', 'pointermove'].forEach((t) => $('#grid').addEventListener(t, (e) => { if (t !== 'pointermove' || Math.abs(e.movementX) + Math.abs(e.movementY) > 3) clearTimeout(lp); }));
function toggleSel(id, card) {
  selected.has(id) ? selected.delete(id) : selected.add(id);
  card.classList.toggle('sel', selected.has(id));
  $('#grid').classList.toggle('selecting', selected.size > 0);
  if (selected.has(id) && !reduced()) card.querySelector('.card-art').animate([{ transform: 'scale(1)' }, { transform: 'scale(0.95)' }, { transform: 'scale(1)' }], { duration: 320, easing: 'cubic-bezier(.2,.9,.3,1.3)' });
  renderSelbar();
}
function cardMenu(btn, id) {
  const d = docById(id);
  openMenu(btn, [
    ...(d.error ? [{ label: 'Try again', icon: 'refresh', run: () => { d.error = null; d.pending = true; renderLibrary(); queueIndex(d.id, true); } }] : [
      { label: 'Read', icon: 'book', run: () => openFree([id]) },
      { label: 'Pace it', icon: 'timer', run: () => openSetup([id]) },
    ]),
    '-',
    { label: 'Rename', icon: 'pen', run: () => renameDoc(d) },
    ...(d.kind === 'images' && !d.local ? [{ label: d.collection ? 'Add photos…' : 'Add photos (make a collection)…', icon: 'plus', run: () => pickPhotos(d) }] : []),
    ...(d.kind === 'images' && d.local ? [{ label: 'Add photos: drop them in its folder', icon: 'folder', run: () => reveal(d.paths[0]) }] : []),
    ...(d.local && d.kind !== 'images' ? [{ label: 'Show in File Explorer', icon: 'folder', run: () => reveal(d.path) }] : []),
    ...(readCount(d) ? [{ label: 'Reset progress', icon: 'reset', run: async () => { d.read = []; d.pos = 0; await store.putDoc(d); renderLibrary(); } }] : []),
    ...(d.local ? [] : [{ label: 'Delete', icon: 'trash', danger: true, run: () => deleteDocs([id]) }]),
  ]);
}
function toggleFolderSel(t, tile) {
  const ids = docsUnder(t.path).map((d) => d.id);
  const all = ids.every((id) => selected.has(id));
  ids.forEach((id) => (all ? selected.delete(id) : selected.add(id)));
  tile.classList.toggle('sel', !all);
  $('#grid').classList.toggle('selecting', selected.size > 0);
  if (!all && !reduced()) tile.querySelector('.card-art').animate([{ transform: 'scale(1)' }, { transform: 'scale(0.95)' }, { transform: 'scale(1)' }], { duration: 320, easing: 'cubic-bezier(.2,.9,.3,1.3)' });
  renderSelbar();
}
function folderMenu(btn, t) {
  const ids = docsUnder(t.path).map((d) => d.id);
  const top = t.path.length === 1 && folders.find((f) => f.path === t.path[0]);
  const fsPath = top ? top.path : [t.path[0], ...t.path.slice(1)].join('\\');
  openMenu(btn, [
    ...(ids.length ? [{ label: 'Read all', icon: 'book', run: () => openFree(readingOrder(ids)) }, { label: 'Pace all', icon: 'timer', run: () => openSetup(readingOrder(ids)) }, '-'] : []),
    ...(isDesktop && folders.some((f) => f.path === t.path[0]) ? [{ label: 'Show in File Explorer', icon: 'folder', run: () => reveal(fsPath) }, { label: 'Check for changes', icon: 'refresh', run: () => syncFolders({ only: t.path[0], announce: 'refresh' }) }] : []),
    ...(top ? ['-', { label: 'Remove from library', icon: 'close', danger: true, run: () => removeFolder(top.path) }] : []),
  ]);
}
function renameDoc(d) {
  d.customName = true;
  const s = openSheet({ title: 'Rename', body: `<div class="group"><div class="row"><input class="field-in text" id="rn" value="${esc(d.name)}" maxlength="140"></div></div>`, foot: '<span class="grow"></span><button class="btn primary" data-save>Save</button>' });
  const inp = $('#rn', s.el); setTimeout(() => { inp.focus(); inp.select(); }, 300);
  const save = async () => { const v = inp.value.trim(); if (v) { d.name = v; await store.putDoc(d); renderLibrary(); } s.close(); };
  s.el.querySelector('[data-save]').onclick = save;
  inp.onkeydown = (e) => { if (e.key === 'Enter') save(); };
}
async function deleteDocs(ids) {
  const gone = ids.map(docById).filter(Boolean);
  for (const id of ids) { await store.deleteDoc(id); selected.delete(id); }
  dropTexts(ids);
  docs = docs.filter((d) => !ids.includes(d.id));
  // a saved session that used them can't continue
  const ses = await store.get('session');
  if (ses && ses.units.some((u) => ids.includes(u.doc))) await store.del('session');
  renderLibrary();
  toast(`Deleted ${gone.length === 1 ? `“${esc(gone[0].name)}”` : plural(gone.length, 'document')}`);
}

// ================= reading speed & history =================
let speed = { wpm: null, slide: null };
async function loadSpeed() { speed = (await store.get('speed')) || speed; }
async function learnSpeed(p) {
  const st = p.stats();
  let changed = false;
  if (st.wpm && st.wpm > 40 && st.wpm < 900) { speed.wpm = Math.round(speed.wpm ? speed.wpm * 0.6 + st.wpm * 0.4 : st.wpm); changed = true; }
  const sl = st.pages.filter((x) => x.done && docById(x.doc)?.kind === 'slides');
  if (sl.length >= 3) { const a = sl.reduce((s, x) => s + x.sec, 0) / sl.length; speed.slide = Math.round(speed.slide ? speed.slide * 0.6 + a * 0.4 : a); changed = true; }
  if (changed) await store.set('speed', speed);
  const hist = (await store.get('history')) || [];
  if (st.work > 60) {
    hist.push({ t: Date.now(), work: Math.round(st.work), pages: st.pages.filter((x) => x.done).length, words: st.words, wpm: st.wpm, docs: [...new Set(p.units.map((u) => u.doc))].map((id) => docById(id)?.name) });
    await store.set('history', hist.slice(-300));
  }
}

// ================= setup: pace a set of documents =================
function unitsFor(cfgDocs) {
  const units = [];
  for (const { id, from, to } of cfgDocs) {
    const d = docById(id);
    if (!d) continue;
    for (let i = from; i <= to; i++) units.push({ doc: id, i, words: d.words[i] || 0, kind: d.kind });
  }
  return units;
}
async function openSetup(ids, { fromReader = null, edit = null } = {}) {
  closeMenu();
  ids = await ensureReady(ids);
  if (!ids.length) return;
  const last = prefs.setup || {};
  const st = {
    docs: ids.map(docById).filter(Boolean).map((d) => {
      // start where you left off when part of it is read
      const read = new Set(d.read || []);
      let from = 0; while (from < d.units - 1 && read.has(from)) from++;
      if (from >= d.units - 1 && read.size >= d.units) from = 0;
      return { id: d.id, from: fromReader?.doc === d.id ? fromReader.i : from, to: d.units - 1 };
    }),
    mode: last.mode || 'total', total: null, per: last.per || 60, by: null,
    profile: last.profile || 'smart', strength: last.strength ?? 70, rebalance: last.rebalance || 'spread',
    method: last.method || 'pomodoro', m: { ...(last.m || {}) }, interrupt: last.interrupt ?? true, recall: last.recall ?? true, breakDocs: last.breakDocs ?? false,
  };
  const units = () => unitsFor(st.docs);
  const suggest = () => suggestTime(units(), speed);
  st.total = Math.max(5 * 60, Math.round(suggest() / 300) * 300 || 1800);
  const d0 = new Date(Date.now() + st.total * 1000 + 15 * 60000); d0.setMinutes(Math.ceil(d0.getMinutes() / 5) * 5, 0, 0);
  st.by = d0.getTime();
  // editing a running session: start from exactly what it's set to
  if (edit) {
    const c = edit.cfg;
    st.docs = [];
    for (const u of edit.units) {
      const q = st.docs.find((x) => x.id === u.doc);
      if (!q) st.docs.push({ id: u.doc, from: u.i, to: u.i }); else { q.from = Math.min(q.from, u.i); q.to = Math.max(q.to, u.i); }
    }
    Object.assign(st, { mode: c.mode, per: c.per, profile: c.profile, strength: c.strength, rebalance: c.rebalance, method: c.method || 'none', interrupt: c.interrupt, recall: c.recall, breakDocs: c.breakDocs });
    st.m = c.flowtime || st.method === 'none' ? {} : c.breakUnit === 'pages' ? { pages: c.breakEvery, short: Math.round(c.breakLen / 60) } : c.breakEvery ? { focus: c.breakEvery, short: Math.round(c.breakLen / 60), long: c.longEvery ? Math.round(c.longLen / 60) : 0, every: c.longEvery || 4 } : {};
    st.total = Math.round(c.total / 60) * 60;
    if (c.by) st.by = c.by;
  }
  const method = () => (st.method === 'custom' ? { id: 'custom', name: 'Custom', ...(prefs.customMethod || { focus: 30, short: 5, long: 15, every: 4 }) } : METHODS.find((m) => m.id === st.method) || METHODS[0]);
  const mvals = () => ({ ...method(), ...st.m });

  const s = openSheet({ title: edit ? 'Edit session' : 'Pace it', wide: false, body: '<div id="su"></div>', foot: `<div class="grow" id="suSum"></div><button class="btn primary big" data-start>${edit ? 'Save' : 'Start'}</button>` });
  const root = $('#su', s.el);
  const timeHtml = () => {
    if (st.mode === 'total') return `<div class="big-time"><div class="stepper"><button data-tadj="-300" aria-label="5 minutes less">−</button></div><div class="readout"><b>${human(st.total).replace(' min', 'm').replace(' h', 'h')}</b><small>${edit ? `for the whole session · ${human(edit.work)} so far` : 'to read it all'}</small></div><div class="stepper"><button data-tadj="300" aria-label="5 minutes more">+</button></div></div>
      <div class="chips">${[15, 30, 45, 60, 90, 120, 180].map((m) => `<button class="chip${st.total === m * 60 ? ' on' : ''}" data-tset="${m * 60}">${m < 60 ? `${m}m` : `${m / 60}h`}</button>`).join('')}<button class="chip accent" data-tset="${Math.round(suggest() / 60) * 60}">Suggested ~${human(suggest())}</button></div>`;
    if (st.mode === 'per') return `<div class="big-time"><div class="stepper"><button data-padj="-5">−</button></div><div class="readout"><b>${clock(st.per, { hours: false })}</b><small>average a page</small></div><div class="stepper"><button data-padj="5">+</button></div></div>
      <div class="chips">${[20, 30, 45, 60, 90, 120, 180, 300].map((v) => `<button class="chip${st.per === v ? ' on' : ''}" data-pset="${v}">${v < 60 ? `${v}s` : `${v / 60}m`}</button>`).join('')}</div>`;
    if (st.mode === 'by') {
      const dt = new Date(st.by);
      const hh = String(dt.getHours()).padStart(2, '0'), mm = String(dt.getMinutes()).padStart(2, '0');
      const left = (st.by - Date.now()) / 1000;
      return `<div class="big-time"><div class="readout"><input type="time" id="byTime" value="${hh}:${mm}"><small>${left > 60 ? `${human(left)} from now${left > 86400 ? '' : ''}` : 'pick a time later today'}</small></div></div>
        <div class="chips">${[30, 60, 90, 120].map((m) => `<button class="chip" data-byin="${m}">in ${m < 60 ? `${m}m` : `${m / 60}h`}</button>`).join('')}</div>`;
    }
    return `<p class="method-blurb" style="padding-top:12px;text-align:center">No deadline. Each page gets a target from your usual pace (${speed.wpm ? `${speed.wpm} words a minute` : 'a careful study pace until it learns yours'}), and your focus cycles run as set below.</p>`;
  };
  const planBars = () => {
    const u = units();
    const w = weightsFor(u, st);
    if (!w.length) return '';
    const n = Math.min(120, w.length);
    const bins = Array.from({ length: n }, (_, b) => { const a = Math.floor((b * w.length) / n), z = Math.max(a + 1, Math.floor(((b + 1) * w.length) / n)); let s2 = 0; for (let i = a; i < z; i++) s2 += w[i]; return s2 / (z - a); });
    const max = Math.max(...bins);
    return `<div class="plan-chart">${bins.map((v) => `<i style="--h:${Math.max(6, (v / max) * 100)}%"></i>`).join('')}</div><div class="plan-cap"><span>first page</span><span>time per page, as planned</span><span>last</span></div>`;
  };
  const methodBar = (m) => {
    if (m.flowtime) return '<span class="m-bar"><i style="flex:5"></i><i class="r" style="flex:1"></i><i style="flex:3"></i><i class="r" style="flex:.6"></i></span>';
    if (m.pages) return '<span class="m-bar"><i style="flex:3"></i><i class="r" style="flex:.6"></i><i style="flex:3"></i><i class="r" style="flex:.6"></i></span>';
    if (!m.focus) return '<span class="m-bar"><i style="flex:1"></i></span>';
    const cyc = m.every || 2; let h = '';
    for (let i = 0; i < Math.min(cyc, 4); i++) { h += `<i style="flex:${m.focus}"></i>`; h += `<i class="${m.long && i === cyc - 1 ? 'L' : 'r'}" style="flex:${m.long && i === cyc - 1 ? m.long : m.short}"></i>`; }
    return `<span class="m-bar">${h}</span>`;
  };
  const pattern = (m) => (m.flowtime ? 'break when you need' : m.pages ? `${m.pages} pages / ${m.short}m` : m.focus ? `${m.focus} / ${m.short}${m.long ? ` · ${m.long}m long` : ''}` : 'free');
  const render = () => {
    const u = units();
    const m = mvals();
    root.innerHTML = `
      <div class="group-t">Reading</div>
      <div class="group" id="queue">${st.docs.map((q, k) => { const d = docById(q.id); return `<div class="qrow" data-q="${k}">
        ${st.docs.length > 1 ? `<span class="q-handle" aria-label="Drag to reorder">${ICON.grip}</span>` : ''}
        <span class="q-thumb">${d.thumb ? `<img src="${d.thumb}" alt="">` : esc(labelOf(d))}</span>
        <div class="q-text"><b>${esc(d.name)}</b><div class="q-range">${unitWord(d)} <input class="field-in" type="number" min="1" max="${d.units}" value="${q.from + 1}" data-from="${k}" inputmode="numeric"> to <input class="field-in" type="number" min="1" max="${d.units}" value="${q.to + 1}" data-to="${k}" inputmode="numeric"> of ${d.units}</div></div>
        ${st.docs.length > 1 ? `<button class="icon-btn" data-qdel="${k}" aria-label="Remove">${ICON.close}</button>` : ''}
      </div>`; }).join('')}<button class="row add-row" data-addq>${ICON.plus}<span>Add documents</span></button></div>
      <div class="group-t">Time</div>
      <div class="group"><div class="row" style="padding-bottom:0"><div class="seg wide" role="group">${[['total', 'Total time'], ['per', 'Per page'], ['by', 'Finish by'], ['open', 'Open']].map(([k, l]) => `<button data-mode="${k}" aria-pressed="${st.mode === k}">${l}</button>`).join('')}</div></div>${timeHtml()}</div>
      <div class="group-t">Study method</div>
      <div class="group"><div class="methods">${[...METHODS, { id: 'custom', name: 'Custom', ...(prefs.customMethod || { focus: 30, short: 5, long: 15, every: 4 }) }].map((x) => `<button class="method${st.method === x.id ? ' on' : ''}" data-method="${x.id}"><b>${esc(x.name)}</b><span class="m-pat">${esc(pattern(x))}</span>${methodBar(x)}</button>`).join('')}</div>
        <p class="method-blurb">${esc(method().blurb || 'Your own rhythm. Change any of the numbers below.')}</p>
        ${m.focus && !m.flowtime ? `<div class="row"><div class="rt"><b>Focus</b></div><div class="stepper"><button data-mv="focus:-5">−</button><output>${m.focus} min</output><button data-mv="focus:5">+</button></div></div>
        <div class="row"><div class="rt"><b>Short break</b></div><div class="stepper"><button data-mv="short:-1">−</button><output>${m.short} min</output><button data-mv="short:1">+</button></div></div>
        <div class="row"><div class="rt"><b>Long break</b><small>${m.long ? `after every ${m.every || 4} focus blocks` : 'off'}</small></div><div class="stepper"><button data-mv="long:-5">−</button><output>${m.long ? `${m.long} min` : 'Off'}</output><button data-mv="long:5">+</button></div></div>
        ${m.long ? `<div class="row"><div class="rt"><b>Long break every</b></div><div class="stepper"><button data-mv="every:-1">−</button><output>${m.every || 4} blocks</output><button data-mv="every:1">+</button></div></div>` : ''}
        <div class="row"><div class="rt"><b>Interrupt me when time’s up</b><small>${st.interrupt ? 'Stops you mid-page; you can take the break, ask for more time, finish the page or skip it' : 'Waits for you to turn the page'}</small></div>${sw('interrupt', st.interrupt)}</div>` : ''}
        ${m.pages ? `<div class="row"><div class="rt"><b>Break every</b></div><div class="stepper"><button data-mv="pages:-1">−</button><output>${m.pages} pages</output><button data-mv="pages:1">+</button></div></div><div class="row"><div class="rt"><b>For</b></div><div class="stepper"><button data-mv="short:-1">−</button><output>${m.short} min</output><button data-mv="short:1">+</button></div></div>` : ''}
        ${m.id !== 'none' ? `<div class="row"><div class="rt"><b>Recall check at breaks</b><small>Before you rest, try to recall what you just read without looking. Retrieval practice is one of the most reliable ways to remember.</small></div>${sw('recall', st.recall)}</div>` : ''}
        ${st.docs.length > 1 ? `<div class="row"><div class="rt"><b>Break between documents</b></div>${sw('breakDocs', st.breakDocs)}</div>` : ''}
      </div>
      <div class="group-t">Pacing</div>
      <div class="group">
        <div class="row stack"><div class="seg wide" role="group">${[['smart', 'Smart'], ['even', 'Even'], ['front', 'Front-load'], ['back', 'Back-load']].map(([k, l]) => `<button data-profile="${k}" aria-pressed="${st.profile === k}">${l}</button>`).join('')}</div>
          <small class="muted">${{ smart: 'Pages with more on them get more time; a title slide gets a moment.', even: 'Every page gets the same time.', front: 'More time early, while you’re fresh.', back: 'More time towards the end, for the harder parts.' }[st.profile]}</small></div>
        ${st.profile === 'smart' ? `<div class="row stack"><div class="rt" style="flex-direction:row;justify-content:space-between"><b>How much content matters</b><output>${st.strength}%</output></div><input type="range" class="sl" min="0" max="100" value="${st.strength}" data-strength></div>` : ''}
        ${planBars()}
      </div>
      <div class="group">
        <div class="row stack"><div class="rt"><b>Finish a page early or late</b></div><div class="seg wide" role="group">${[['spread', 'Share it out'], ['carry', 'Next page'], ['fixed', 'Keep plan']].map(([k, l]) => `<button data-reb="${k}" aria-pressed="${st.rebalance === k}">${l}</button>`).join('')}</div>
        <small class="muted">${{ spread: 'Time you save or lose is shared over every page still to go.', carry: 'The difference rolls onto the very next page.', fixed: 'Each page keeps its time; you just see if you’re ahead or behind.' }[st.rebalance]}</small></div>
      </div>`;
    root.querySelectorAll('input[type=range]').forEach(fillTrack);
    summary();
  };
  const summary = () => {
    const u = units();
    const n = u.length;
    let tot = st.mode === 'total' ? st.total : st.mode === 'per' ? st.per * n : st.mode === 'by' ? (st.by - Date.now()) / 1000 : suggest();
    if (edit && st.mode !== 'by') tot = Math.max(60, tot - edit.work);
    const p = new Pacer(u.length ? u : [{ doc: 'x', i: 0, words: 0 }], cfg());
    const brk = p.plannedBreaks(n, st.mode === 'by' ? tot : tot, st.docs.length);
    if (st.mode === 'by') tot = Math.max(60, tot - brk);
    const end = st.mode === 'by' ? st.by : Date.now() + (tot + brk) * 1000;
    const left = edit ? u.filter((x) => !edit.done.has(`${x.doc}:${x.i}`)).length || n : n;
    $('#suSum', s.el).innerHTML = n ? `<b>${edit ? `${left} of ${n} pages left` : `${n} pages`} · ${clock(tot / left, { hours: false })} a page</b><br>${st.mode === 'open' ? 'open-ended' : `done around ${timeOfDay(end)}`}${brk ? ` · ${human(brk)} of breaks` : ''}` : 'Nothing to read';
    s.el.querySelector('[data-start]').disabled = !n || (st.mode === 'by' && st.by - Date.now() < 120000);
  };
  const cfg = () => {
    const mc = methodCfg(mvals(), {});
    return { mode: st.mode, total: st.mode === 'open' ? suggest() : st.total, per: st.per, by: st.mode === 'by' ? st.by : null, profile: st.profile, strength: st.strength, rebalance: st.rebalance, ...mc, interrupt: mc.flowtime ? false : st.interrupt, recall: st.recall, breakDocs: st.breakDocs };
  };
  root.addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    const ds = b.dataset;
    if (ds.mode) st.mode = ds.mode;
    else if (ds.tadj) st.total = Math.max(300, Math.min(16 * 3600, st.total + Number(ds.tadj)));
    else if (ds.tset) st.total = Number(ds.tset);
    else if (ds.padj) st.per = Math.max(5, Math.min(1800, st.per + Number(ds.padj) * (st.per >= 120 ? 6 : st.per >= 60 ? 3 : 1)));
    else if (ds.pset) st.per = Number(ds.pset);
    else if (ds.byin) { st.by = Date.now() + Number(ds.byin) * 60000; }
    else if (ds.method) { st.method = ds.method; st.m = {}; }
    else if (ds.mv) {
      const [k, dv] = ds.mv.split(':');
      const m = mvals();
      const lim = { focus: [5, 180], short: [1, 60], long: [0, 60], every: [2, 8], pages: [1, 100] }[k];
      st.m[k] = Math.max(lim[0], Math.min(lim[1], (m[k] || 0) + Number(dv)));
      if (k === 'long' && st.m.long && !m.every) st.m.every = 4;
      // edits to a preset become your Custom rhythm
      prefs.customMethod = { focus: mvals().focus || 30, short: mvals().short || 5, long: mvals().long || 0, every: mvals().every || 4, blurb: undefined };
      if (st.method !== 'custom' && !method().pages && !method().flowtime) { const keep = { ...mvals() }; st.method = 'custom'; st.m = {}; prefs.customMethod = { focus: keep.focus, short: keep.short, long: keep.long, every: keep.every || 4 }; }
    }
    else if (ds.profile) st.profile = ds.profile;
    else if (ds.reb) st.rebalance = ds.reb;
    else if (ds.qdel) st.docs.splice(Number(ds.qdel), 1);
    else if (b.hasAttribute('data-addq')) { pickDocs(st.docs.map((q) => q.id), (add) => { for (const d of add) st.docs.push({ id: d.id, from: 0, to: d.units - 1 }); const y = s.body.scrollTop; render(); s.body.scrollTop = y; }); return; }
    else return;
    const y = s.body.scrollTop;
    render();
    s.body.scrollTop = y;
  });
  root.addEventListener('input', (e) => {
    const t = e.target;
    if (t.dataset.strength != null && t.hasAttribute('data-strength')) { st.strength = Number(t.value); fillTrack(t); t.closest('.row').querySelector('output').textContent = `${st.strength}%`; const pc = root.querySelector('.plan-chart'); if (pc) pc.outerHTML = planBars().split('<div class="plan-cap">')[0]; summary(); return; }
    if (t.id === 'byTime') { const [h, m] = t.value.split(':').map(Number); const d = new Date(); d.setHours(h, m, 0, 0); if (d.getTime() < Date.now() + 60000) d.setDate(d.getDate() + 1); st.by = d.getTime(); t.nextElementSibling.textContent = `${human((st.by - Date.now()) / 1000)} from now`; summary(); return; }
    if (t.dataset.k) { st[t.dataset.k] = t.checked; const y = s.body.scrollTop; render(); s.body.scrollTop = y; return; }
  });
  root.addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.from != null || t.dataset.to != null) {
      const k = Number(t.dataset.from ?? t.dataset.to);
      const q = st.docs[k], d = docById(q.id);
      const v = Math.max(1, Math.min(d.units, Math.round(Number(t.value) || 1))) - 1;
      if (t.dataset.from != null) q.from = Math.min(v, q.to); else q.to = Math.max(v, q.from);
      const y = s.body.scrollTop; render(); s.body.scrollTop = y;
    }
  });
  // drag to reorder documents
  root.addEventListener('pointerdown', (e) => {
    const h = e.target.closest('.q-handle');
    if (!h) return;
    e.preventDefault();
    const row = h.closest('.qrow');
    const rows = $$('.qrow', root);
    const from = Number(row.dataset.q);
    const hgt = row.offsetHeight;
    let to = from;
    const y0 = e.clientY;
    row.classList.add('dragging');
    h.setPointerCapture(e.pointerId);
    const move = (ev) => {
      const dy = ev.clientY - y0;
      row.style.transform = `translateY(${dy}px)`;
      to = Math.max(0, Math.min(rows.length - 1, from + Math.round(dy / hgt)));
      rows.forEach((r, i) => { if (r === row) return; const shift = from < to && i > from && i <= to ? -hgt : from > to && i < from && i >= to ? hgt : 0; r.style.transition = 'transform .25s cubic-bezier(.2,.9,.3,1)'; r.style.transform = shift ? `translateY(${shift}px)` : ''; });
    };
    const up = () => {
      h.removeEventListener('pointermove', move); h.removeEventListener('pointerup', up); h.removeEventListener('pointercancel', up);
      const [it] = st.docs.splice(from, 1); st.docs.splice(to, 0, it);
      render();
    };
    h.addEventListener('pointermove', move); h.addEventListener('pointerup', up); h.addEventListener('pointercancel', up);
  });
  s.el.querySelector('[data-start]').onclick = () => {
    const u = units();
    if (!u.length) return;
    prefs.setup = { mode: st.mode, per: st.per, profile: st.profile, strength: st.strength, rebalance: st.rebalance, method: st.method, m: st.m, interrupt: st.interrupt, recall: st.recall, breakDocs: st.breakDocs };
    savePrefs();
    s.close();
    if (edit) { edit.onSave(u, cfg()); return; }
    const pacer = new Pacer(u, cfg());
    startSession(u, pacer);
  };
  render();
}

// add documents from the library to a session
function pickDocs(exclude, done) {
  const pool = docs.filter((d) => !exclude.includes(d.id) && !d.pending && !d.error && d.units);
  const order = readingOrder(pool.map((d) => d.id)).map(docById);
  const where = (d) => (d.dir || []).map(folderLabel).join(' › ');
  const s = openSheet({
    title: 'Add documents',
    body: order.length ? `<div class="group">${order.map((d) => `<button class="row pick-row multi" data-id="${d.id}"><span class="pick-dot"></span><div class="rt"><b>${esc(d.name)}</b><small>${esc([where(d), `${d.units} ${unitWord(d, d.units)}`].filter(Boolean).join(' · '))}</small></div></button>`).join('')}</div>` : '<p class="muted" style="margin:6px">Everything in your library is already in this session.</p>',
    foot: '<span class="grow" id="pdCount"></span><button class="btn primary" data-ok disabled>Add</button>',
  });
  const picked = new Set();
  s.body.onclick = (e) => {
    const r = e.target.closest('[data-id]');
    if (!r) return;
    picked.has(r.dataset.id) ? picked.delete(r.dataset.id) : picked.add(r.dataset.id);
    r.classList.toggle('on', picked.has(r.dataset.id));
    s.el.querySelector('[data-ok]').disabled = !picked.size;
    $('#pdCount', s.el).textContent = picked.size ? plural(picked.size, 'document') : '';
  };
  s.el.querySelector('[data-ok]').onclick = () => { s.close(); done(order.filter((d) => picked.has(d.id))); };
}
function editSession() {
  const p = R.pacer;
  if (!p) return;
  p.tick();
  const ids = [...new Set(p.units.map((u) => u.doc))];
  const done = new Set(p.units.filter((_, j) => p.s.done[j]).map((u) => `${u.doc}:${u.i}`));
  openSetup(ids, { edit: { cfg: { ...p.cfg, total: p.cfg.mode === 'total' ? p.s.budget + p.s.added : p.cfg.total }, units: p.units, work: p.s.work, done, onSave: applyEdit } });
}
async function applyEdit(units, cfg) {
  const old = R.pacer;
  if (!old) return;
  old.tick();
  const p = carryOver(old, new Pacer(units, cfg));
  const flagged = new Set([...R.flags].map((k) => old.units[k]).filter(Boolean).map((u) => `${u.doc}:${u.i}`));
  R.session.units = units;
  await openReader(units, p, p.s.cur, null, { quiet: true });
  units.forEach((u, k) => { if (flagged.has(`${u.doc}:${u.i}`)) { R.flags.add(k); R.viewer.setFlag(k, true); } });
  renderTop();
  saveSession();
  toast(`${ICON.check}Session updated`, { ms: 1800 });
}

// ================= the reader =================
const R = { viewer: null, units: [], pacer: null, k: 0, flags: new Set(), docs: new Map(), raf: 0, session: null, recall: [], lastText: 0, awayPaused: false, veil: null };
const reader = $('#reader');

Pacer.restore = (ses) => { const p = new Pacer(ses.units, ses.pacer.cfg, ses.pacer.s); return p; };

async function openFree(ids, fromCard) {
  ids = await ensureReady(ids);
  const ds = ids.map(docById).filter(Boolean);
  if (!ds.length) return;
  const units = [];
  for (const d of ds) for (let i = 0; i < d.units; i++) units.push({ doc: d.id, i, words: d.words[i] || 0, kind: d.kind });
  const start = ds.length === 1 ? Math.min(ds[0].pos || 0, ds[0].units - 1) : 0;
  await openReader(units, null, start, fromCard);
}
async function startSession(units, pacer) {
  await store.del('session');
  R.recall = [];
  const s = { id: Date.now(), units, pacer: pacer.toJSON(), flags: [], recall: [], created: Date.now() };
  R.session = s;
  await openReader(units, pacer, 0);
  pacer.start();
  saveSession();
  toast(`${ICON.timer} ${pacer.cfg.mode === 'open' ? 'Pacing started' : `${human(pacer.snapshot().remaining)} for ${plural(units.length, 'page')}`}${pacer.cfg.method !== 'none' ? ` · ${esc(METHODS.find((m) => m.id === pacer.cfg.method)?.name || 'Custom')}` : ''}`, { ms: 2600 });
}
async function resumeSession() {
  const ses = await store.get('session');
  if (!ses) return;
  if (ses.units.some((u) => !docById(u.doc))) { toast('Some documents in that session were deleted'); await store.del('session'); renderHero(); return; }
  const p = Pacer.restore(ses);
  R.session = ses;
  R.recall = ses.recall || [];
  await openReader(ses.units, p, p.s.cur);
  (ses.flags || []).forEach((k) => { R.flags.add(k); R.viewer.setFlag(k, true); });
  p.resume();
  if (p.s.brk) showBreak();
}
async function saveSession() {
  if (!R.pacer || !R.session) return;
  R.session.pacer = R.pacer.toJSON();
  R.session.flags = [...R.flags];
  R.session.recall = R.recall;
  try { await store.set('session', R.session); } catch {}
}

async function openReader(units, pacer, start, fromEl, { quiet = false } = {}) {
  closeMenu();
  R.units = units; R.pacer = pacer; R.flags = new Set(); R.search = null; $('#hitbar').hidden = true;
  R.docs = new Map(units.map((u) => [u.doc, docById(u.doc)]));
  R.k = start;
  $('#library').inert = true;
  reader.hidden = false;
  reader.classList.remove('immersive');
  reader.classList.toggle('paged-mode', prefs.view !== 'scroll');
  // zoom in from the card that was tapped
  if (!reduced() && !quiet) {
    const r = fromEl?.getBoundingClientRect?.();
    if (r) {
      reader.style.transformOrigin = `${r.left + r.width / 2}px ${r.top + r.height / 2}px`;
      reader.animate([{ transform: `scale(${Math.max(0.2, r.width / innerWidth)})`, opacity: 0, borderRadius: '30px' }, { transform: 'none', opacity: 1, borderRadius: '0' }], { duration: 480, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1)' });
    } else reader.animate([{ opacity: 0, transform: 'scale(1.03)' }, { opacity: 1, transform: 'none' }], { duration: 380, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1)' });
  }
  R.viewer?.destroy();
  const v = new Viewer($('#stage'), {
    docs: R.docs, units, prefs,
    onPage: (k, dir) => moved(k, dir, true),
    onTap: (e) => tapStage(e),
    onSwipe: (dir) => { const j = R.k + dir; if (j < 0 || j >= units.length) { if (dir > 0 && R.pacer) nextPage(); return false; } go(j); return true; },
  });
  v.onZoom = (z) => showZoom(z);
  R.viewer = v;
  await v.primeNotes();
  v.show(start, 0);
  renderTop();
  renderHud(true);
  cancelAnimationFrame(R.raf);
  loop();
  keepAwake(!!pacer);
  $('#stage').focus({ preventScroll: true });
  if (!quiet) history.pushState({ reader: 1 }, '');
}
function exitReader(fromPop) {
  if (reader.hidden || R.exiting) return;
  R.exiting = true;
  cancelAnimationFrame(R.raf);
  if (R.pacer && !R.pacer.s.finished) { R.pacer.tick(); saveSession(); }
  closeVeil(true);
  $('.nudge')?.remove();
  R.search = null; $('#hitbar').hidden = true;
  const out = reader.animate([{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'scale(0.97)' }], { duration: reduced() ? 1 : 240, easing: 'ease-in' });
  out.onfinish = () => { reader.hidden = true; R.exiting = false; R.viewer?.destroy(); R.viewer = null; $('#stage').innerHTML = ''; $('#hud').innerHTML = ''; };
  $('#library').inert = false;
  keepAwake(false);
  R.pacer = null;
  loadLibrary();
  if (!fromPop && history.state?.reader) history.back();
}
addEventListener('popstate', () => { if (sheets.length) { sheets[sheets.length - 1].close(); history.pushState({ reader: 1 }, ''); return; } if (!reader.hidden) exitReader(true); });

function go(k) {
  if (k < 0 || k >= R.units.length || k === R.k) return;
  const dir = k > R.k ? 1 : -1;
  R.viewer.show(k, Math.abs(k - R.k) > 1 ? 0 : dir);
  moved(k, dir, false);
}
function moved(k, dir, fromScroll) {
  const from = R.k;
  R.k = k;
  if (R.pacer) {
    const r = R.pacer.go(k);
    if (r?.break) { chime('due'); buzz([30, 60, 30]); showBreak(); }
  }
  // remember what's been read and where you are
  const u = R.units[k], d = R.docs.get(u.doc);
  const fu = R.units[from];
  if (fu && dir > 0) markRead(fu);
  if (d) { d.pos = u.i; d.opened = Date.now(); clearTimeout(R._dp); R._dp = setTimeout(() => store.putDoc(d), 600); }
  renderTop();
  renderHud();
  if (R.search) paintHitbar();
  saveSession();
}
function markRead(u) {
  const d = R.docs.get(u.doc);
  if (!d) return;
  d.read ||= [];
  if (!d.read.includes(u.i)) { d.read.push(u.i); store.putDoc(d); }
}
function nextPage() {
  if (R.viewer.mode === 'paged' && R.k < R.units.length - 1) { go(R.k + 1); return; }
  if (R.k < R.units.length - 1) { go(R.k + 1); return; }
  // the very end
  markRead(R.units[R.k]);
  if (!R.pacer) { toast(`${ICON.check} That’s the end`, { actions: [{ label: 'Library', run: () => exitReader() }] }); return; }
  const r = R.pacer.next();
  if (r.finished) finishSession();
  else if (r.unread != null) {
    const n = R.pacer.left();
    toast(`${plural(n, 'page')} still unread`, { ms: 6000, actions: [{ label: 'Go there', primary: true, run: () => go(r.unread) }, { label: 'Finish', run: () => { R.pacer.finish(); finishSession(); } }] });
  }
}
const prevPage = () => go(R.k - 1);
function tapStage(e) {
  if (R.veil) return;
  const x = e.clientX / innerWidth;
  if (prefs.tapEdges && R.viewer.mode === 'paged' && e.pointerType !== 'mouse') {
    if (x < 0.22) { prevPage(); return; }
    if (x > 0.78) { nextPage(); return; }
  }
  if (e.pointerType === 'mouse') return;
  reader.classList.toggle('immersive');
}
function renderTop() {
  const u = R.units[R.k], d = R.docs.get(u.doc);
  $('#rdName').textContent = d.name;
  const docIds = [...R.docs.keys()];
  const head = d.heads?.[u.i];
  $('#rdSub').textContent = `${d.kind === 'slides' ? 'Slide' : 'Page'} ${u.i + 1} of ${d.units}${docIds.length > 1 ? ` · document ${docIds.indexOf(u.doc) + 1} of ${docIds.length}` : ''}${head && head !== d.name ? ` · ${head}` : ''}`;
  const fb = $('#flagBtn');
  fb.classList.toggle('on', R.flags.has(R.k));
  // progress, a segment per document
  const prog = $('#rdProg');
  const segs = docIds.map((id) => {
    const ks = R.units.map((x, j) => (x.doc === id ? j : -1)).filter((j) => j >= 0);
    const done = R.pacer ? ks.filter((j) => R.pacer.s.done[j]).length : ks.filter((j) => j < R.k).length;
    return { id, n: ks.length, p: (done / ks.length) * 100 };
  });
  if (prog.childElementCount !== segs.length) prog.innerHTML = segs.map((s) => `<span class="seg-doc" style="--n:${s.n}"><i></i></span>`).join('');
  segs.forEach((s, i) => { const el = prog.children[i]; el.style.setProperty('--p', `${s.p}%`); el.classList.toggle('cur', s.id === u.doc); });
}

// ---- the pacer bar ----
const RING_R = 23, RING_C = 2 * Math.PI * RING_R, FO_R = 28.5, FO_C = 2 * Math.PI * FO_R;
function renderHud(fresh) {
  const hud = $('#hud');
  hud.classList.toggle('compact', !!prefs.hudCompact);
  if (!R.pacer) {
    const u = R.units[R.k], d = R.docs.get(u.doc);
    const left = R.units.length - R.k - 1;
    hud.classList.add('free-hud');
    hud.innerHTML = `<div class="hud-stats"><div class="hs hs-main"><b>${R.k + 1} / ${R.units.length}</b><small>${left ? `${plural(left, d.kind === 'slides' ? 'slide' : 'page')} to go` : 'last page'}</small></div></div>
      <div class="hud-ctrl"><button class="icon-btn prev" data-act="prev" aria-label="Previous">${ICON.prev}</button><button class="btn primary" data-act="paceThis">${ICON.timer}<span>Pace it</span></button><button class="icon-btn" data-act="next" aria-label="Next">${ICON.next}</button></div>`;
    if (fresh) { hud.classList.remove('in'); void hud.offsetWidth; hud.classList.add('in'); }
    return;
  }
  hud.classList.remove('free-hud');
  if (fresh || !hud.querySelector('.ring')) {
    hud.innerHTML = `<div class="ring" data-act="ring" title="Pause or resume (P)"><svg viewBox="0 0 64 64"><circle class="tr" cx="32" cy="32" r="${RING_R}"/><circle class="ar" cx="32" cy="32" r="${RING_R}" stroke-dasharray="${RING_C}"/><circle class="fo" cx="32" cy="32" r="${FO_R}" stroke-dasharray="${FO_C}" stroke-dashoffset="${FO_C}"/></svg><b></b></div>
      <div class="hud-stats"><button class="hs hs-main" data-act="stats"><b></b><small></small></button><button class="hs hs-two" data-act="overview"><b></b><small></small></button><button class="hs hs-third" data-act="stats"><b></b><small></small></button></div>
      <div class="hud-ctrl"><button class="icon-btn prev" data-act="prev" aria-label="Previous page">${ICON.prev}</button><button class="icon-btn play" data-act="pause" aria-label="Pause"></button><button class="icon-btn" data-act="next" aria-label="Next page">${ICON.next}</button></div>`;
    hud.classList.remove('in'); void hud.offsetWidth; hud.classList.add('in');
    R._playState = null;
  }
  paintHud(true);
}
function paintHud(text) {
  const p = R.pacer;
  if (!p) return;
  const sn = p.snapshot();
  const hud = $('#hud');
  const ring = hud.querySelector('.ring');
  if (!ring) return;
  // page ring: drains as the page's time runs out
  const ar = ring.querySelector('.ar');
  let frac, label, cls = '';
  if (sn.review) { frac = 1; label = clock(sn.pageSpent, { hours: false }); cls = 'review'; }
  else if (sn.pageLeft >= 0) { frac = sn.pageAlloc ? sn.pageLeft / sn.pageAlloc : 0; label = clock(sn.pageLeft, { hours: false }); cls = frac < 0.2 ? 'warn' : ''; }
  else { frac = Math.min(1, -sn.pageLeft / Math.max(10, sn.pageAlloc)); label = `+${clock(-sn.pageLeft, { hours: false })}`; cls = 'over'; }
  ar.style.strokeDashoffset = String(RING_C * (1 - Math.max(0, Math.min(1, frac))));
  ring.className = `ring ${cls}${sn.paused ? ' paused' : ''}`;
  // outer arc: the focus block
  const fo = ring.querySelector('.fo');
  if (sn.focusLeft != null) { const tot = p.cfg.breakEvery * 60 + (p.s.snooze || 0); fo.style.strokeDashoffset = String(FO_C * (1 - Math.max(0, Math.min(1, sn.focusLeft / tot)))); fo.style.opacity = 1; }
  else fo.style.opacity = 0;
  if (!text) return;
  ring.querySelector('b').textContent = sn.paused ? 'II' : label;
  const [a, b, c] = hud.querySelectorAll('.hs');
  const dots = p.cfg.longEvery ? `<span class="focus-dots">${Array.from({ length: p.cfg.longEvery }, (_, i) => `<i class="${i < sn.cycles % p.cfg.longEvery ? 'on' : ''}"></i>`).join('')}</span>` : '';
  if (sn.open) { a.querySelector('b').textContent = clock(sn.work); a.querySelector('small').innerHTML = `${sn.focusLeft != null ? `focus ${clock(Math.max(0, sn.focusLeft), { hours: false })}` : 'reading time'}${dots}`; }
  else {
    a.querySelector('b').innerHTML = sn.remaining >= 0 ? clock(sn.remaining) : `<span class="bad">${clock(sn.remaining)}</span>`;
    a.querySelector('small').innerHTML = sn.focusLeft != null ? `focus ${clock(Math.max(0, sn.focusLeft), { hours: false })}${dots}` : `${sn.remaining >= 0 ? 'left' : 'over'}${dots}`;
  }
  b.querySelector('b').textContent = sn.left ? `${sn.left} to go` : 'All read';
  b.querySelector('small').textContent = `${sn.review ? 'reviewing · ' : ''}${sn.docs.length > 1 ? `${plural(sn.docsLeft, 'doc')} · ` : ''}${sn.doneN} of ${sn.n} done`;
  const ahead = Math.round(sn.delta / 60);
  c.querySelector('b').textContent = timeOfDay(sn.predicted);
  c.querySelector('small').innerHTML = sn.trust < 0.15 || sn.open ? 'expected finish' : Math.abs(ahead) < 1 ? '<span class="good">right on pace</span>' : ahead > 0 ? `<span class="good">${ahead} min ahead</span>` : `<span class="${ahead < -5 ? 'bad' : 'warn'}">${-ahead} min behind</span>`;
  const play = hud.querySelector('.play');
  const st = sn.paused ? 'play' : 'pause';
  if (R._playState !== st) { play.innerHTML = ICON[st]; play.setAttribute('aria-label', sn.paused ? 'Resume' : 'Pause'); R._playState = st; }
}
function loop() {
  const tickOnce = () => {
    const p = R.pacer;
    if (p && !p.s.finished) {
      p.tick();
      const now = performance.now();
      const text = now - R.lastText > 220;
      if (text) R.lastText = now;
      paintHud(text);
      const sn = text ? p.snapshot() : null;
      if (sn) {
        if (sn.due && !R.veil) showDue();
        if (sn.brk && !R.veil) showBreak();
        if (R.veil?.kind === 'break') paintBreak(sn);
        // a gentle heads-up a minute before a focus block ends
        if (prefs.nudge && sn.focusLeft != null && p.cfg.interrupt && sn.focusLeft < 60 && sn.focusLeft > 0 && !R._nudged && !sn.paused) { R._nudged = true; nudge(`Break in a minute · wrap up your thought`); chime('soft'); }
        if (sn.focusLeft != null && sn.focusLeft > 60) R._nudged = false;
        if (sn.over && !p.s.overNotified) { p.s.overNotified = true; outOfTime(sn); }
        if (now - (R._saved || 0) > 4000) { R._saved = now; saveSession(); }
      }
    }
    R.raf = requestAnimationFrame(tickOnce);
  };
  R.raf = requestAnimationFrame(tickOnce);
}
function nudge(text) {
  $('.nudge')?.remove();
  const n = document.createElement('div');
  n.className = 'nudge glass-menu';
  n.textContent = text;
  reader.append(n);
  setTimeout(() => { n.classList.add('out'); setTimeout(() => n.remove(), 320); }, 5200);
}
function outOfTime(sn) {
  chime('end'); buzz([40, 80, 40]);
  toast(`⏰ Time’s up · ${plural(sn.left, 'page')} to go`, { ms: 12000, actions: [{ label: '+5 min', run: () => addTime(300) }, { label: '+15', run: () => addTime(900) }, { label: 'Keep going', primary: true }] });
}
function addTime(sec) { R.pacer?.addTime(sec); toast(`${ICON.plus} Added ${human(sec)}`, { ms: 1800 }); saveSession(); paintHud(true); }
function togglePause() {
  const p = R.pacer;
  if (!p) return;
  if (p.s.paused) { p.resume(); R.awayPaused = false; } else p.pause();
  buzz(8);
  const ring = $('#hud .ring');
  ring?.animate([{ transform: 'scale(1)' }, { transform: 'scale(0.86)' }, { transform: 'scale(1)' }], { duration: 360, easing: 'cubic-bezier(.2,.9,.3,1.4)' });
  paintHud(true);
  saveSession();
}

// ---- interruptions: the focus block is done ----
function confetti(colors) {
  if (reduced()) return;
  const c = document.createElement('div');
  c.className = 'confetti';
  const cols = colors || ['#e2552c', '#f5b400', '#2f9e44', '#2f54eb', '#d6457b', '#0e9aa7'];
  c.innerHTML = Array.from({ length: 46 }, () => {
    const a = Math.random() * Math.PI * 2, r = 160 + Math.random() * 260;
    return `<i style="--c:${cols[Math.floor(Math.random() * cols.length)]};--x:${Math.cos(a) * r}px;--y:${Math.sin(a) * r - 80}px;--r:${Math.random() * 720 - 360}deg;--d:${0.75 + Math.random() * 0.25}s"></i>`;
  }).join('');
  document.body.append(c);
  setTimeout(() => c.remove(), 2800);
}
function openVeil(kind, html) {
  closeVeil(true);
  const v = document.createElement('div');
  v.className = 'veil';
  v.innerHTML = `<div class="veil-card">${html}</div>`;
  document.body.append(v);
  R.veil = { kind, el: v };
  closeMenu();
  return v;
}
function closeVeil(fast) {
  const v = R.veil;
  if (!v) return;
  R.veil = null;
  if (fast) { v.el.remove(); return; }
  v.el.classList.add('out');
  setTimeout(() => v.el.remove(), 400);
}
function showDue() {
  const p = R.pacer, sn = p.snapshot();
  const due = sn.due;
  const pages = p.s.log.filter((e) => e.type === 'page' && e.t > (p.s.lastBreakAt || p.s.started || 0)).length;
  const focus = Math.round((p.cfg.breakEvery * 60 + (p.s.snooze || 0)) / 60);
  chime('due'); buzz([30, 70, 30, 70, 60]);
  const v = openVeil('due', `
    <div class="burst done" style="--bc:#e2552c"><span class="ripple"></span><span class="ripple"></span><span class="ripple"></span>
      <svg class="dial" viewBox="0 0 160 160"><circle class="tr" cx="80" cy="80" r="70"/><circle class="ar" cx="80" cy="80" r="70"/></svg>
      <span class="core">${ICON.check}</span></div>
    <h2>${due.long ? 'Set complete. Long break!' : 'Focus block done!'}</h2>
    <p>${focus} minutes of focus${pages ? `, ${plural(pages, 'page')} read` : ''}. ${due.long ? `That’s ${p.cfg.longEvery} blocks in a row. You’ve earned a proper rest.` : 'Step away for a moment; it helps it stick.'}</p>
    <div class="veil-actions">
      <button class="btn primary full" data-v="take">${ICON.coffee}<span>Take ${human(due.len)} break</span></button>
      <button class="btn" data-v="s5">+5 minutes</button><button class="btn" data-v="s10">+10 minutes</button>
      <button class="btn" data-v="page">Finish this page</button><button class="btn" data-v="skip">Skip break</button>
    </div>
    <div class="veil-more"><button class="chip" data-v="s2">+2</button><button class="chip" data-v="s15">+15</button><button class="chip" data-v="s25">+25 min</button></div>`);
  confetti();
  v.onclick = (e) => {
    const b = e.target.closest('[data-v]');
    if (!b) return;
    const a = b.dataset.v;
    if (a === 'take') { p.acceptDue(); closeVeil(true); showBreak(); }
    else if (a === 'page') { p.deferDue(); closeVeil(); toast('Break when you turn the page', { ms: 2200 }); }
    else if (a === 'skip') { p.skipDue(); closeVeil(); toast('Skipped. Next focus block started', { ms: 2200 }); }
    else if (a[0] === 's') { const m = Number(a.slice(1)); p.snoozeDue(m * 60); closeVeil(); toast(`${m} more minutes`, { ms: 1800 }); }
    saveSession();
  };
}

// ---- the break itself ----
const TIPS = ['Stand up and stretch your back and shoulders.', 'Look at something 20 feet away for 20 seconds to rest your eyes.', 'Drink some water.', 'Take a short walk, even just around the room.', 'Close your eyes and take a few slow breaths.', 'Roll your neck and loosen your jaw.', 'Step away from screens; your brain keeps consolidating while you rest.'];
function showBreak() {
  const p = R.pacer, sn = p.snapshot();
  if (!sn.brk) return;
  const since = (sn.brk.since || []).map((k) => R.units[k]).filter(Boolean);
  const heads = [...new Set(since.map((u) => R.docs.get(u.doc)?.heads?.[u.i]).filter((h) => h && h.length > 2))].slice(-6);
  const recall = p.cfg.recall && since.length && sn.brk.why !== 'manual';
  const C = 2 * Math.PI * 104;
  const v = openVeil('break', `
    <div class="breathe"><span class="orb"></span><span class="orb o2"></span><svg viewBox="0 0 220 220"><circle class="tr" cx="110" cy="110" r="104"/><circle class="ar" cx="110" cy="110" r="104" stroke-dasharray="${C}"/></svg>
      <div class="bt"><b id="bkLeft">${clock(sn.brk.left, { hours: false })}</b><span id="bkBreath">Breathe in</span></div></div>
    <h2 id="bkTitle">${sn.brk.long ? 'Long break' : sn.brk.why === 'doc' ? 'One document down' : 'Break time'}</h2>
    <p class="tip">${esc(TIPS[Math.floor(Math.random() * TIPS.length)])}</p>
    ${recall ? `<div class="recall"><small>Recall check</small><p>Without looking back, what were the key ideas from the last ${plural(since.length, 'page')}?</p>${heads.length ? `<ul>${heads.map((h) => `<li>${esc(h)}</li>`).join('')}</ul>` : ''}<textarea id="bkNote" placeholder="Jot down what you remember (saved with this session)"></textarea></div>` : ''}
    <div class="veil-actions">
      <button class="btn" data-b="m1">+1 minute</button><button class="btn" data-b="m5">+5 minutes</button>
      <button class="btn primary full" data-b="end" id="bkEnd">I’m ready, back to it</button>
    </div>`);
  R.veil.C = C;
  R.veil.startT = performance.now();
  v.onclick = (e) => {
    const b = e.target.closest('[data-b]');
    if (!b) return;
    if (b.dataset.b === 'm1') p.extendBreak(60);
    else if (b.dataset.b === 'm5') p.extendBreak(300);
    else if (b.dataset.b === 'end') endBreakNow();
    paintBreak(p.snapshot());
  };
}
function paintBreak(sn) {
  const v = R.veil?.el;
  if (!v || !sn.brk) { if (R.veil?.kind === 'break' && !sn.brk) closeVeil(); return; }
  const left = sn.brk.left;
  $('#bkLeft', v).textContent = left > 0 ? clock(left, { hours: false }) : '0:00';
  v.querySelector('.breathe .ar').style.strokeDashoffset = String(R.veil.C * (1 - Math.max(0, left) / sn.brk.len));
  // 10 s breathing: 4 in, 1 hold, 5 out (matches the orb)
  const ph = ((performance.now() - R.veil.startT) / 1000) % 10;
  const word = ph < 4 ? 'Breathe in' : ph < 5 ? 'Hold' : 'Breathe out';
  const bw = $('#bkBreath', v); if (bw.textContent !== word && left > 0) bw.textContent = word;
  if (left <= 0 && !R.veil.ready) {
    R.veil.ready = true;
    chime('end'); buzz([40, 60, 40]);
    $('#bkTitle', v).textContent = 'Break’s over';
    bw.textContent = 'Ready when you are';
    const end = $('#bkEnd', v); end.animate([{ transform: 'scale(1)' }, { transform: 'scale(1.05)' }, { transform: 'scale(1)' }], { duration: 700, iterations: 3 });
  }
}
function endBreakNow() {
  const p = R.pacer;
  const note = $('#bkNote')?.value.trim();
  if (note) R.recall.push({ t: Date.now(), note, pages: p.s.brk?.since || [] });
  p.endBreak();
  closeVeil();
  toast('Back to it. Focus block started', { ms: 1800 });
  saveSession();
}

// ---- finished ----
async function finishSession() {
  const p = R.pacer;
  p.finish();
  const st = p.stats();
  chime('finish'); buzz([30, 50, 30, 50, 90]);
  await learnSpeed(p);
  await store.del('session');
  const sn = p.snapshot();
  const delta = Math.round((p.cfg.mode === 'by' ? (p.cfg.by - Date.now()) : (p.s.budget + p.s.added - p.s.work) * 1000) / 60000);
  const v = openVeil('finish', `
    <div class="burst done fin"><span class="ripple"></span><span class="ripple"></span><span class="ripple"></span>
      <svg class="dial" viewBox="0 0 160 160"><circle class="tr" cx="80" cy="80" r="70"/><circle class="ar" cx="80" cy="80" r="70"/></svg><span class="core">${ICON.check}</span></div>
    <h2>All done!</h2>
    <p>${plural(sn.n, 'page')} in ${human(st.work)}${st.breaks ? ` with ${plural(st.breaks, 'break')}` : ''}.${p.cfg.mode !== 'open' ? ` ${delta >= 1 ? `${delta} min ahead of plan.` : delta <= -1 ? `${-delta} min over plan.` : 'Right on time.'}` : ''}${st.wpm ? ` You read at about ${st.wpm} words a minute.` : ''}</p>
    <div class="veil-actions"><button class="btn" data-f="stats">${ICON.chart}<span>Session stats</span></button><button class="btn primary" data-f="lib">Library</button></div>`);
  confetti(['#2f9e44', '#5dd17c', '#f5b400', '#2f54eb', '#e2552c']);
  v.onclick = (e) => {
    const b = e.target.closest('[data-f]');
    if (!b) return;
    if (b.dataset.f === 'lib') { closeVeil(); exitReader(); }
    else { closeVeil(); openStats(p); }
  };
}
async function endSessionAsk() {
  closeMenu();
  const p = R.pacer;
  if (!p) return;
  const s = openSheet({ title: 'End this session?', body: `<p class="muted" style="margin:0 6px 8px">${plural(p.left(), 'page')} still to go. Your time so far is kept in your history.</p>`, foot: '<span class="grow"></span><button class="btn" data-x="keep">Keep going</button><button class="btn primary" data-x="end">End session</button>' });
  s.el.querySelector('[data-x="keep"]').onclick = () => s.close();
  s.el.querySelector('[data-x="end"]').onclick = async () => { s.close(); p.finish(); await learnSpeed(p); await store.del('session'); R.session = null; openStats(p, true); };
}

// ---- overview of every page ----
function openOverview() {
  closeMenu();
  const p = R.pacer;
  const groups = [];
  R.units.forEach((u, k) => { let g = groups[groups.length - 1]; if (!g || g.doc !== u.doc) { g = { doc: u.doc, ks: [] }; groups.push(g); } g.ks.push(k); });
  const body = groups.map((g) => {
    const d = R.docs.get(g.doc);
    const done = p ? g.ks.filter((k) => p.s.done[k]).length : 0;
    return `<div class="ov-doc"><b>${esc(d.name)}</b><small>${p ? `${done} of ${g.ks.length} read` : `${g.ks.length} ${unitWord(d)}`}</small></div>
      <div class="ov-grid">${g.ks.map((k) => {
        const u = R.units[k];
        const spent = p ? p.s.spent[k] : 0;
        return `<button class="ov-cell${k === R.k ? ' cur' : ''}${p?.s.done[k] ? ' done' : ''}${R.flags.has(k) ? ' flag' : ''}" data-k="${k}"><span class="ov-art" style="--ar:${d.kind === 'flow' ? 0.77 : u.ar || d.aspect || 0.77}"></span><small>${u.i + 1}${spent > 1 ? ` · <b>${clock(spent, { hours: false })}</b>` : ''}</small></button>`;
      }).join('')}</div>`;
  }).join('');
  const flagged = [...R.flags].sort((a, b) => a - b);
  const s = openSheet({ title: 'Pages', wide: true, body: `${flagged.length ? `<div class="group-t">Flagged</div><div class="chips" style="justify-content:flex-start;padding:0 2px 10px">${flagged.map((k) => `<button class="chip accent" data-k="${k}">★ ${esc(R.docs.get(R.units[k].doc).name.slice(0, 24))} · ${R.units[k].i + 1}</button>`).join('')}</div>` : ''}${body}` });
  s.body.onclick = (e) => { const c = e.target.closest('[data-k]'); if (c) { s.close(); go(Number(c.dataset.k)); } };
  const io = new IntersectionObserver((ents) => {
    for (const en of ents) {
      if (!en.isIntersecting) continue;
      io.unobserve(en.target);
      const k = Number(en.target.closest('[data-k]').dataset.k);
      R.viewer?.thumb(k, 110).then(({ el }) => { en.target.append(el); el.animate?.([{ opacity: 0 }, { opacity: 1 }], { duration: 200 }); }).catch(() => {});
    }
  }, { root: s.body, rootMargin: '200px' });
  $$('.ov-art', s.body).forEach((a) => io.observe(a));
  requestAnimationFrame(() => s.body.querySelector('.ov-cell.cur')?.scrollIntoView({ block: 'center' }));
}

// ---- stats ----
function openStats(p = R.pacer, ended = false) {
  closeMenu();
  if (!p) return;
  const st = p.stats();
  const pages = st.pages.filter((x) => x.done || x.sec > 2);
  const max = Math.max(1, ...pages.map((x) => Math.max(x.sec, x.planned || 0)));
  const recall = (R.session?.recall || R.recall || []);
  const name = (j) => { const u = p.units[j]; const d = docById(u.doc); return `${esc(d?.name?.slice(0, 30) || '')} · ${u.i + 1}`; };
  const body = `
    <div class="stat-grid">
      <div class="stat"><b>${human(st.work)}</b><span>reading time</span></div>
      <div class="stat"><b>${pages.filter((x) => x.done).length}/${p.units.length}</b><span>pages done</span></div>
      <div class="stat"><b>${st.avg ? clock(st.avg, { hours: false }) : '–'}</b><span>average a page</span></div>
      <div class="stat"><b>${st.wpm ? st.wpm : '–'}</b><span>words a minute</span></div>
      <div class="stat"><b>${st.breaks}</b><span>${st.breaks === 1 ? 'break' : 'breaks'} · ${human(st.breakTotal)}</span></div>
      <div class="stat"><b>${p.s.cycles || 0}</b><span>focus blocks</span></div>
    </div>
    ${pages.length ? `<div class="group"><div class="bars">${pages.map((x, k) => `<span class="b" title="${name(x.j).replace(/"/g, '')}: ${clock(x.sec, { hours: false })}${x.planned ? ` of ${clock(x.planned, { hours: false })}` : ''}"><i class="${x.planned && x.sec > x.planned * 1.05 ? 'over' : ''}" style="--h:${(x.sec / max) * 100}%;--k:${Math.min(k, 80)}"></i>${x.planned ? `<u style="--p:${(x.planned / max) * 100}%"></u>` : ''}</span>`).join('')}</div>
      <div class="bars-cap"><span><i></i>time on page</span><span><i class="o"></i>over its time</span><span><i class="p"></i>planned</span></div></div>` : ''}
    ${st.fastest ? `<div class="group"><div class="row"><div class="rt"><b>Quickest</b><small>${name(st.fastest.j)}</small></div><output>${clock(st.fastest.sec, { hours: false })}</output></div><div class="row"><div class="rt"><b>Longest</b><small>${name(st.slowest.j)}</small></div><output>${clock(st.slowest.sec, { hours: false })}</output></div></div>` : ''}
    ${recall.length ? `<div class="group-t">What you recalled</div><div class="group">${recall.map((r) => `<div class="row"><div class="rt"><small>${timeOfDay(r.t)}</small><b style="font-weight:500;white-space:pre-wrap">${esc(r.note)}</b></div></div>`).join('')}</div>` : ''}
    <div class="group-t">Log</div>
    <div class="group logl">${p.s.log.filter((e) => ['page', 'break', 'add', 'snooze', 'skip', 'pause', 'resume'].includes(e.type)).slice(-80).reverse().map((e) => `<div class="row"><span class="muted" style="width:70px">${timeOfDay(e.t)}</span><span class="rt">${e.type === 'page' ? `Read ${name(e.i)}` : e.type === 'break' ? `${e.why === 'manual' ? 'Took' : 'Started'} a ${human(e.len)} break` : e.type === 'add' ? `Added ${human(e.sec)}` : e.type === 'snooze' ? `Asked for ${human(e.sec)} more` : e.type === 'skip' ? 'Skipped a break' : e.type === 'pause' ? 'Paused' : 'Resumed'}</span>${e.type === 'page' ? `<output>${clock(e.sec, { hours: false })}</output>` : ''}</div>`).join('') || '<div class="row muted">Nothing yet</div>'}</div>`;
  const s = openSheet({ title: ended ? 'Session summary' : 'This session', wide: true, body, foot: ended ? '<span class="grow"></span><button class="btn primary" data-x="lib">Back to library</button>' : `<span class="grow"></span><button class="btn danger" data-x="end">End session</button>`, onClose: () => { if (ended && !reader.hidden) exitReader(); } });
  s.el.querySelector('[data-x]')?.addEventListener('click', (e) => { const x = e.currentTarget.dataset.x; s.close(); if (x === 'end') endSessionAsk(); });
}

// ---- more ----
function moreMenu(btn) {
  const p = R.pacer;
  openMenu(btn, [
    ...(p ? [
      { label: p.s.paused ? 'Resume' : 'Pause', icon: p.s.paused ? 'play' : 'pause', key: 'P', run: togglePause },
      { label: 'Take a break now', icon: 'coffee', key: 'B', run: () => { p.takeBreakNow(); showBreak(); } },
      { label: 'Edit session…', icon: 'pen', key: 'E', run: editSession },
      { label: 'Add 5 minutes', icon: 'plus', run: () => addTime(300) },
      { label: 'Session stats', icon: 'chart', key: 'S', run: () => openStats() },
      '-',
    ] : [{ label: 'Pace these pages', icon: 'timer', run: () => paceThis() }, '-']),
    { label: prefs.view === 'scroll' ? 'Page by page' : 'Continuous scroll', icon: 'view', key: 'V', run: toggleView },
    { label: prefs.fit === 'width' ? 'Fit whole page' : 'Fit to width', icon: 'zoom', run: () => { prefs.fit = prefs.fit === 'width' ? 'page' : 'width'; savePrefs(); R.viewer.relayout(); } },
    { label: prefs.hudCompact ? 'Full pacer bar' : 'Compact pacer bar', icon: 'timer', run: () => { prefs.hudCompact = !prefs.hudCompact; savePrefs(); renderHud(); } },
    ...([...R.docs.values()].some((d) => d.kind === 'slides') ? [{ label: prefs.notes ? 'Hide speaker notes' : 'Show speaker notes', icon: 'notes', run: () => { prefs.notes = !prefs.notes; savePrefs(); R.viewer.lru.clear(); const k = R.k; R.viewer.k = -1; R.viewer.show(k, 0); } }] : []),
    { label: 'Settings', icon: 'gear', run: openSettings },
    ...(p ? ['-', { label: 'End session', icon: 'stop', danger: true, run: endSessionAsk }] : []),
  ]);
}
function toggleView() { prefs.view = prefs.view === 'scroll' ? 'paged' : 'scroll'; savePrefs(); reader.classList.toggle('paged-mode', prefs.view !== 'scroll'); R.viewer.setMode(); toast(prefs.view === 'scroll' ? 'Continuous scroll' : 'Page by page', { ms: 1400 }); }
function paceThis() {
  const ids = [...R.docs.keys()];
  const u = R.units[R.k];
  exitReader();
  setTimeout(() => openSetup(ids, { fromReader: u }), 260);
}
function showZoom(z) {
  const pill = $('#zoomPill');
  pill.hidden = false;
  pill.innerHTML = `<button class="icon-btn" data-z="-1" aria-label="Zoom out">${ICON.minus}</button><output>${Math.round(z * 100)}%</output><button class="icon-btn" data-z="1" aria-label="Zoom in">${ICON.plus}</button><button class="icon-btn" data-z="0" aria-label="Reset zoom">${ICON.reset}</button>`;
  clearTimeout(R._zh); R._zh = setTimeout(() => { pill.hidden = true; }, 3500);
}
$('#zoomPill').onclick = (e) => { const b = e.target.closest('[data-z]'); if (!b) return; const d = Number(b.dataset.z); showZoom(R.viewer.setZoom(d === 0 ? 1 : R.viewer.zoom * (d > 0 ? 1.2 : 1 / 1.2))); };
function toggleFlag() {
  const on = !R.flags.has(R.k);
  on ? R.flags.add(R.k) : R.flags.delete(R.k);
  R.viewer.setFlag(R.k, on);
  const b = $('#flagBtn');
  b.classList.toggle('on', on);
  b.animate([{ transform: 'scale(1) rotate(0)' }, { transform: 'scale(1.35) rotate(-14deg)' }, { transform: 'scale(1)' }], { duration: 420, easing: 'cubic-bezier(.2,.9,.3,1.4)' });
  buzz(8);
  saveSession();
}


// ================= search =================
// Every document's words are indexed by page when it's read in (headings and captions marked), and searched
// in a worker with Findly's engine: Any word, All words, Exact phrase, Near and Boolean, with match case,
// whole words, and where on the page to look.
let worker = null;
const searchWorker = () => { if (!worker) { worker = new Worker(new URL('./search-worker.js', import.meta.url), { type: 'module' }); worker.onmessage = onWorker; } return worker; };
const SV = { qid: 0, last: null, open: false, timer: 0, expanded: new Set() };
const sv = $('#search');
const qIn = $('#q');
// documents still waiting for their text index (imported before search existed)
const textJob = { queue: [], busy: false };
async function backfillTexts() {
  const have = new Set(await store.textIds());
  for (const d of docs) if (!have.has(d.id) && !d.pending && !d.error && d.units && !textJob.queue.includes(d.id)) textJob.queue.push(d.id);
  if (textJob.busy) return;
  textJob.busy = true;
  while (textJob.queue.length) {
    const d = docById(textJob.queue.shift());
    if (!d || d.pending) continue;
    try { await store.putText(d.id, await buildText(d)); } catch (e) { console.warn('index', d.name, e); }
    if (SV.open) renderSearch();
    await new Promise((r) => setTimeout(r, 30));
  }
  textJob.busy = false;
  if (SV.open) renderSearch();
}
const dropTexts = (ids) => { try { worker?.postMessage({ t: 'drop', ids }); } catch {} };

function scopeOptions() {
  const o = [];
  if (!reader.hidden) {
    const ids = [...R.docs.keys()];
    if (ids.length === 1) o.push(['doc', 'This document']);
    else { o.push(['doc', 'This document']); o.push(['session', R.pacer ? 'This session' : 'These documents']); }
  } else if (cwd.length) o.push(['folder', `In ${folderLabel(cwd[cwd.length - 1])}`]);
  o.push(['all', 'Everything']);
  return o;
}
function scopeIds() {
  const ok = (d) => d && !d.pending && !d.error && d.units;
  const sc = SV.scope;
  if (sc === 'doc' && !reader.hidden) return [R.units[R.k].doc];
  if (sc === 'session' && !reader.hidden) return [...R.docs.keys()];
  if (sc === 'folder') return readingOrder(docsUnder(cwd).filter(ok).map((d) => d.id));
  return readingOrder(docs.filter(ok).map((d) => d.id));
}
function openSearch(prefill) {
  closeMenu();
  const opts = scopeOptions();
  if (!SV.open || !opts.some(([k]) => k === SV.scope)) SV.scope = opts.some(([k]) => k === prefs.sScope) ? prefs.sScope : opts[0][0];
  SV.open = true;
  sv.hidden = false;
  document.body.classList.add('searching');
  if (!reduced()) sv.animate([{ opacity: 0, transform: 'translateY(-8px)' }, { opacity: 1, transform: 'none' }], { duration: 260, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1)' });
  if (typeof prefill === 'string') qIn.value = prefill;
  renderScope();
  renderMode();
  runSearch();
  setTimeout(() => { qIn.focus(); qIn.select(); }, 30);
  backfillTexts();
}
function closeSearch() {
  if (!SV.open) return;
  SV.open = false;
  closeMenu();
  document.body.classList.remove('searching');
  const a = sv.animate([{ opacity: 1 }, { opacity: 0, transform: 'translateY(-6px)' }], { duration: reduced() ? 1 : 180, easing: 'ease-in' });
  a.onfinish = () => { if (!SV.open) sv.hidden = true; };
  qIn.blur();
  if (!reader.hidden) $('#stage').focus({ preventScroll: true });
}
function renderScope() {
  const opts = scopeOptions();
  $('#svScope').innerHTML = opts.length > 1 ? `<div class="seg small" role="group">${opts.map(([k, l]) => `<button data-scope="${k}" aria-pressed="${SV.scope === k}">${esc(l)}</button>`).join('')}</div>` : '';
}
function renderMode() {
  const flags = [prefs.matchCase ? 'Aa' : '', prefs.wholeWord ? 'Whole words' : '', prefs.sMode === 'near' ? `within ${prefs.near}` : ''].filter(Boolean);
  const sub = [WHERE[prefs.where]?.short, ...flags].filter(Boolean).join(' · ');
  $('#svMode').innerHTML = `<span><b>${MODES[prefs.sMode].label}</b>${sub ? `<small>${esc(sub)}</small>` : ''}</span>${ICON.down}`;
  qIn.placeholder = { any: 'Any of these words', all: 'Search your documents', phrase: 'An exact phrase', near: 'Words close together', boolean: 'e.g. (stall OR spin) AND recovery NOT training' }[prefs.sMode];
}
const NEAR_STEPS = [2, 3, 5, 8, 12, 20, 40, 80];
const PICK_HINT = { any: 'Any of the words', all: 'Every word, anywhere on the page', phrase: 'The words together, in order', near: 'All the words close together', boolean: 'AND, OR, NOT and ( )' };
function pickerHtml() {
  const opts = Object.keys(MODES).map((m) => `<button class="pk-opt${prefs.sMode === m ? ' on' : ''}" data-mode="${m}" role="option" aria-selected="${prefs.sMode === m}"><span class="pk-txt"><b>${MODES[m].label}</b><span>${PICK_HINT[m]}</span></span>${ICON.check}</button>`).join('');
  const near = prefs.sMode === 'near' ? `<div class="pk-near"><span>Within</span><button class="pk-step" data-near="-1" aria-label="Closer">${ICON.minus}</button><b>${prefs.near}</b><button class="pk-step" data-near="1" aria-label="Further apart">${ICON.plus}</button><span>words</span></div>` : '';
  return `<div class="pk-opts" role="listbox" aria-label="Search type">${opts}</div>${near}<div class="pk-togs"><button class="pk-tog" data-tog="case" aria-pressed="${prefs.matchCase}"><i>Aa</i>Match case</button><button class="pk-tog" data-tog="word" aria-pressed="${prefs.wholeWord}"><i><u>ab</u></i>Whole words</button></div><div class="pk-where"><span>Look in</span><div class="pk-wseg">${Object.entries(WHERE).map(([k, w]) => `<button type="button" data-where="${k}" aria-pressed="${(prefs.where || 'any') === k}">${w.label}</button>`).join('')}</div></div>`;
}
function openPicker(anchor) {
  if (menuEl?.classList.contains('picker')) { closeMenu(); return; }
  closeMenu();
  const m = document.createElement('div');
  m.className = 'menu picker glass-menu';
  m.innerHTML = pickerHtml();
  document.body.append(m);
  const place = () => {
    const r = anchor.getBoundingClientRect();
    const mw = m.offsetWidth;
    const left = Math.min(innerWidth - mw - 8, Math.max(8, r.right - mw));
    m.style.left = `${left}px`; m.style.top = `${r.bottom + 6}px`;
    m.style.setProperty('--ox', `${r.left + r.width / 2 - left}px top`);
  };
  place();
  m.onclick = (e) => {
    e.stopPropagation();
    const o = e.target.closest('.pk-opt');
    if (o) { prefs.sMode = o.dataset.mode; savePrefs(); renderMode(); runSearch(); if (o.dataset.mode === 'near') { m.innerHTML = pickerHtml(); return; } closeMenu(); qIn.focus(); return; }
    const t = e.target.closest('[data-tog]');
    if (t) { if (t.dataset.tog === 'case') prefs.matchCase = !prefs.matchCase; else prefs.wholeWord = !prefs.wholeWord; }
    const wh = e.target.closest('[data-where]');
    if (wh) prefs.where = wh.dataset.where;
    const n = e.target.closest('[data-near]');
    if (n) { const cur = Number(prefs.near) || 8; prefs.near = Number(n.dataset.near) > 0 ? (NEAR_STEPS.find((x) => x > cur) ?? cur) : ([...NEAR_STEPS].reverse().find((x) => x < cur) ?? cur); }
    if (!t && !wh && !n) return;
    savePrefs(); renderMode(); runSearch();
    m.innerHTML = pickerHtml();
  };
  menuEl = m;
  setTimeout(() => document.addEventListener('pointerdown', function off(e) { if (menuEl === m && !m.contains(e.target) && !anchor.contains(e.target)) closeMenu(); document.removeEventListener('pointerdown', off, true); }, true));
}
function runSearch() {
  clearTimeout(SV.timer);
  SV.timer = setTimeout(() => {
    const q = qIn.value;
    $('.sv-clear').hidden = !q;
    const qid = ++SV.qid;
    SV.pending = qid;
    if (!q.trim()) { SV.last = null; renderSearch(); return; }
    const ids = scopeIds();
    searchWorker().postMessage({ t: 'q', qid, query: q, opts: { mode: prefs.sMode, matchCase: prefs.matchCase, wholeWord: prefs.wholeWord, near: prefs.near, where: prefs.where }, ids });
    // a spinner only if it takes a moment
    setTimeout(() => { if (SV.pending === qid) $('#svBody').classList.add('busy'); }, 160);
  }, 130);
}
function onWorker(e) {
  const d = e.data;
  if (d.qid !== SV.qid) return;
  SV.pending = 0;
  $('#svBody').classList.remove('busy');
  SV.last = { ...d, q: qIn.value };
  SV.expanded.clear();
  renderSearch();
}
const pageWord = (d) => (d.kind === 'slides' ? 'Slide' : d.kind === 'images' ? 'Photo' : 'p.');
function hlText(s, hl) {
  let out = '', at = 0;
  for (const [a, b, c] of hl) { out += esc(s.slice(at, a)) + `<mark class="m c${c < 0 ? 0 : c}">${esc(s.slice(a, b))}</mark>`; at = b; }
  return out + esc(s.slice(at));
}
function renderSearch() {
  const body = $('#svBody');
  const r = SV.last;
  const ids = scopeIds();
  const total = ids.reduce((a, id) => a + (docById(id)?.units || 0), 0);
  const waiting = textJob.queue.length + (textJob.busy ? 1 : 0) + docs.filter((d) => d.pending).length;
  if (!r || r.empty) {
    const tips = {
      any: [['stall spin', 'Pages with either word, or both'], ['"fuel pump"', 'Quotes keep words together'], ['hydraul*', 'A star matches any ending'], ['-checklist', 'A minus leaves out pages with that word']],
      all: [['engine failure', 'Pages with both words, anywhere on the page'], ['VFR minimums class E', 'Each word you add narrows it down'], ['"fuel pump" magneto', 'Quotes keep words together'], ['hydraul*', 'A star matches any ending'], ['-checklist', 'A minus leaves out pages with that word']],
      phrase: [['engine failure after takeoff', 'These exact words, in this order'], ['go-around', 'Any dash or curly quote style matches'], ['hydraul* system', 'A star still matches any ending']],
      near: [['carburetor ice', `Both words within ${prefs.near} words of each other`], ['spin recovery rudder', 'Every word close together, in any order'], ['"angle of attack" stall', 'Quotes keep a phrase together']],
      boolean: [['(stall OR spin) NOT training', 'AND, OR, NOT and parentheses, in capitals'], ['VOR OR NDB OR GPS', 'Any one of them'], ['(METAR OR TAF) AND (IFR OR LIFR)', 'Group as deep as you like']],
    }[prefs.sMode];
    body.innerHTML = `<div class="sv-idle"><h3>Search ${plural(ids.length, 'document')} · ${plural(total, 'page')}</h3>${waiting ? `<p class="muted"><span class="spin"></span> Indexing ${plural(waiting, 'document')} for search…</p>` : ''}
      <div class="sv-tips">${tips.map(([c, t]) => `<button data-try="${esc(c)}"><code>${esc(c)}</code><span>${esc(t)}</span></button>`).join('')}<div><span class="keys"><kbd>Enter</kbd></span><span>Open the first result</span></div></div></div>`;
    return;
  }
  if (r.error) { body.innerHTML = `<div class="sv-idle"><h3>Can’t search that yet</h3><p class="muted">${esc(r.error)}</p></div>`; return; }
  if (!r.results.length) {
    const fixes = [];
    if (prefs.sMode === 'all' && /\s/.test(r.q.trim())) fixes.push('<button class="link" data-fix="any">Try Any word</button>');
    if (prefs.sMode === 'near' || prefs.sMode === 'phrase') fixes.push('<button class="link" data-fix="all">Try All words</button>');
    if (prefs.where !== 'any') fixes.push('<button class="link" data-fix="where">Look anywhere on the page</button>');
    if (prefs.matchCase) fixes.push('<button class="link" data-fix="case">Turn off match case</button>');
    if (prefs.wholeWord) fixes.push('<button class="link" data-fix="word">Turn off whole words</button>');
    if (SV.scope !== 'all') fixes.push('<button class="link" data-fix="scope">Search everything</button>');
    body.innerHTML = `<div class="sv-idle"><h3>No matches</h3><p class="muted">Nothing found in ${plural(r.searched, 'document')}${waiting ? ` (${waiting} still indexing)` : ''}.</p><div class="fixes">${fixes.join('')}</div></div>`;
    return;
  }
  const head = `<div class="sv-sum">${r.hits.toLocaleString()} ${r.hits === 1 ? 'match' : 'matches'} on ${plural(r.pagesHit, 'page')} in ${plural(r.results.length, 'document')}${waiting ? ` <span class="muted">· ${waiting} still indexing</span>` : ''}</div>`;
  body.innerHTML = head + r.results.slice(0, 120).map((g, gi) => {
    const d = docById(g.id);
    if (!d) return '';
    const open = SV.expanded.has(g.id);
    const shown = open ? g.pages : g.pages.slice(0, 4);
    const where = (d.dir || []).map(folderLabel).join(' › ');
    return `<section class="sr-doc" style="--i:${Math.min(gi, 12)}"><header><span class="card-badge ${d.kind}">${esc(labelOf(d))}</span><b>${esc(d.name)}</b><small>${esc([where, `${plural(g.pages.length, d.kind === 'slides' ? 'slide' : 'page')}`, `${g.count} ${g.count === 1 ? 'match' : 'matches'}`].filter(Boolean).join(' · '))}</small></header>
      ${shown.map((p) => `<button class="sr-page" data-id="${g.id}" data-n="${p.n}"><span class="sr-pn">${pageWord(d)} ${p.n + 1}</span><span class="sr-txt">${p.h ? `<i>${esc(p.h)}</i>` : ''}${p.s != null ? hlText(p.s, p.hl) : '<span class="muted">Match on this page</span>'}</span><span class="sr-c">${p.c}</span></button>`).join('')}
      ${g.pages.length > 4 && !open ? `<button class="sr-more" data-more="${g.id}">Show all ${g.pages.length} ${d.kind === 'slides' ? 'slides' : 'pages'}</button>` : ''}</section>`;
  }).join('');
}
// open a result: inside the running session when the page is part of it, otherwise that document on its own
async function openResult(id, n) {
  const r = SV.last;
  const hitsFor = new Map((r?.results || []).map((g) => [g.id, g.pages.map((p) => p.n)]));
  const m = compile(r.q, { mode: prefs.sMode, matchCase: prefs.matchCase, wholeWord: prefs.wholeWord, near: prefs.near });
  closeSearch();
  let k = -1;
  if (!reader.hidden) k = R.units.findIndex((u) => u.doc === id && u.i === n);
  if (k < 0) {
    const d = docById(id);
    const units = Array.from({ length: d.units }, (_, i) => ({ doc: id, i, words: d.words[i] || 0, kind: d.kind }));
    // swap the document in place; a running session is saved and waits on the library page
    const inReader = !reader.hidden;
    if (inReader && R.pacer && !R.pacer.s.finished) { R.pacer.tick(); await saveSession(); toast('Session saved · resume it from the library', { ms: 2600 }); }
    await openReader(units, null, n, null, { quiet: inReader });
    k = n;
  } else go(k);
  setHits(m, r.q, hitsFor);
  if (R.viewer.k !== k) go(k);
}
// the hit bar in the reader: which match page you're on, and the next and previous ones
function setHits(m, q, hitsFor) {
  if (!m || m.error) return;
  R.search = { m, q, ks: R.units.map((u, k) => (hitsFor.get(u.doc)?.includes(u.i) ? k : -1)).filter((k) => k >= 0) };
  R.viewer.setHighlight(m);
  paintHitbar(true);
}
function clearHits() { if (!R.search) return; R.search = null; R.viewer?.setHighlight(null); const hb = $('#hitbar'); hb.hidden = true; }
function paintHitbar(fresh) {
  const hb = $('#hitbar');
  const s = R.search;
  if (!s) { hb.hidden = true; return; }
  const at = s.ks.indexOf(R.k);
  const before = s.ks.filter((k) => k < R.k).length;
  hb.hidden = false;
  hb.innerHTML = `${ICON.search}<button class="hb-q" data-hb="edit" title="Change the search">${esc(s.q)}</button><b>${at >= 0 ? `${at + 1} of ${s.ks.length}` : `${before}/${s.ks.length}`}</b><button class="icon-btn" data-hb="prev" aria-label="Previous match page"${before ? '' : ' disabled'}>${ICON.up}</button><button class="icon-btn" data-hb="next" aria-label="Next match page"${s.ks.some((k) => k > R.k) ? '' : ' disabled'}>${ICON.down}</button><button class="icon-btn" data-hb="close" aria-label="Clear search">${ICON.close}</button>`;
  if (fresh && !reduced()) hb.animate([{ opacity: 0, transform: 'translate(-50%, -10px) scale(0.96)' }, { opacity: 1, transform: 'translate(-50%, 0)' }], { duration: 320, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1.1)' });
}
$('#hitbar').onclick = (e) => {
  const b = e.target.closest('[data-hb]');
  if (!b || !R.search) return;
  const s = R.search;
  if (b.dataset.hb === 'close') clearHits();
  else if (b.dataset.hb === 'edit') openSearch(s.q);
  else if (b.dataset.hb === 'next') { const k = s.ks.find((x) => x > R.k); if (k != null) go(k); }
  else { const k = [...s.ks].reverse().find((x) => x < R.k); if (k != null) go(k); }
};
qIn.addEventListener('input', runSearch);
qIn.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); if (qIn.value) { qIn.value = ''; runSearch(); } else closeSearch(); }
  else if (e.key === 'Enter') { e.preventDefault(); const f = $('#svBody .sr-page'); if (f) openResult(f.dataset.id, Number(f.dataset.n)); }
  else if (e.key === 'ArrowDown') { e.preventDefault(); $('#svBody .sr-page')?.focus(); }
});
sv.addEventListener('click', (e) => {
  const b = e.target.closest('button');
  if (!b) return;
  const a = b.dataset;
  if (a.sv === 'close') closeSearch();
  else if (a.sv === 'clear') { qIn.value = ''; runSearch(); qIn.focus(); }
  else if (a.sv === 'mode') openPicker(b);
  else if (a.scope) { SV.scope = a.scope; prefs.sScope = a.scope; savePrefs(); renderScope(); runSearch(); }
  else if (a.try) { qIn.value = a.try; runSearch(); qIn.focus(); }
  else if (a.more) { SV.expanded.add(a.more); const y = $('#svBody').scrollTop; renderSearch(); $('#svBody').scrollTop = y; }
  else if (a.fix) {
    if (a.fix === 'any' || a.fix === 'all') prefs.sMode = a.fix;
    else if (a.fix === 'where') prefs.where = 'any';
    else if (a.fix === 'case') prefs.matchCase = false;
    else if (a.fix === 'word') prefs.wholeWord = false;
    else if (a.fix === 'scope') { SV.scope = 'all'; renderScope(); }
    savePrefs(); renderMode(); runSearch();
  } else if (a.id) openResult(a.id, Number(a.n));
});
$('#svBody').addEventListener('keydown', (e) => {
  const cur = e.target.closest('.sr-page');
  if (!cur) return;
  const all = $$('#svBody .sr-page');
  const i = all.indexOf(cur);
  if (e.key === 'ArrowDown') { e.preventDefault(); all[i + 1]?.focus(); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); (all[i - 1] || qIn).focus(); }
});

// ================= settings =================
function openSettings() {
  closeMenu();
  const seg = (k, opts) => `<div class="seg" role="group">${opts.map(([v, l]) => `<button data-set="${k}" data-v="${v}" aria-pressed="${String(prefs[k]) === String(v)}">${l}</button>`).join('')}</div>`;
  const render = async () => {
    const est = await store.estimate();
    return `
    <div class="group-t">Reading</div>
    <div class="group">
      <div class="row"><div class="rt"><b>Layout</b></div>${seg('view', [['paged', 'Pages'], ['scroll', 'Scroll']])}</div>
      <div class="row"><div class="rt"><b>Fit</b></div>${seg('fit', [['page', 'Whole page'], ['width', 'Width']])}</div>
      <div class="row"><div class="rt"><b>Page turn</b></div>${seg('turn', [['slide', 'Slide'], ['stack', 'Stack'], ['fade', 'Fade'], ['none', 'None']])}</div>
      <div class="row"><div class="rt"><b>Tap the edges to turn pages</b><small>Tap the middle to hide the bars</small></div>${sw('tapEdges', prefs.tapEdges)}</div>
      <div class="row"><div class="rt"><b>Speaker notes under slides</b></div>${sw('notes', prefs.notes)}</div>
    </div>
    <div class="group-t">Text documents</div>
    <div class="group">
      <div class="row"><div class="rt"><b>Font</b></div>${seg('flowFont', [['serif', 'Serif'], ['sans', 'Sans'], ['plain', 'Wide'], ['mono', 'Mono']])}</div>
      <div class="row stack"><div class="rt" style="flex-direction:row;justify-content:space-between"><b>Size</b><output>${prefs.fontSize}px</output></div><input type="range" class="sl" min="13" max="28" value="${prefs.fontSize}" data-r="fontSize"></div>
      <div class="row stack"><div class="rt" style="flex-direction:row;justify-content:space-between"><b>Line spacing</b><output>${prefs.lineH}</output></div><input type="range" class="sl" min="1.2" max="2.2" step="0.05" value="${prefs.lineH}" data-r="lineH"></div>
      <div class="row stack"><div class="rt" style="flex-direction:row;justify-content:space-between"><b>Page width</b><output>${prefs.flowW}px</output></div><input type="range" class="sl" min="480" max="1100" step="20" value="${prefs.flowW}" data-r="flowW"></div>
    </div>
    <div class="group-t">Pacer</div>
    <div class="group">
      <div class="row"><div class="rt"><b>Pause when I leave</b><small>Switching apps or tabs pauses the clock</small></div>${sw('autoPause', prefs.autoPause)}</div>
      <div class="row"><div class="rt"><b>Heads-up before a break</b><small>A quiet note a minute before the focus block ends</small></div>${sw('nudge', prefs.nudge)}</div>
      <div class="row"><div class="rt"><b>Sounds</b></div>${sw('sound', prefs.sound)}</div>
      <div class="row"><div class="rt"><b>Vibration</b></div>${sw('haptics', prefs.haptics)}</div>
      <div class="row"><div class="rt"><b>Keep the screen on</b><small>While a session is running</small></div>${sw('wake', prefs.wake)}</div>
      <div class="row"><div class="rt"><b>Compact pacer bar</b></div>${sw('hudCompact', prefs.hudCompact)}</div>
    </div>
    ${isDesktop ? `<div class="group-t">Folders</div>
    <div class="group">${folders.map((f) => `<div class="row"><div class="rt"><b>${esc(f.name)}</b><small>${esc(f.path)}${f.missing ? ' · not found' : ` · ${plural(docs.filter((d) => d.root === f.path).length, 'document')}`}</small></div><button class="icon-btn" data-fx="show" data-p="${esc(f.path)}" aria-label="Show in File Explorer">${ICON.folder}</button><button class="icon-btn" data-fx="rm" data-p="${esc(f.path)}" aria-label="Remove from library">${ICON.close}</button></div>`).join('')}
      <div class="row"><div class="rt"><b>Add a folder</b><small>Documents are read straight from it, subfolders and all. Changes show up when you come back to the app.</small></div><button class="btn" data-x="addFolder">Add</button></div>
      ${folders.length ? '<div class="row"><div class="rt"><b>Check for changes now</b></div><button class="btn" data-x="refresh">Check</button></div>' : ''}
    </div>` : ''}
    <div class="group-t">Look</div>
    <div class="group"><div class="row"><div class="rt"><b>Theme, paper and glass</b><small>Colours, page textures, see-through and blur</small></div><button class="btn" data-x="themes">Open</button></div></div>
    <div class="group-t">Your data</div>
    <div class="group">
      <div class="row"><div class="rt"><b>Stored on this device</b><small>${docs.length} documents${est?.usage ? ` · ${(est.usage / 1048576).toFixed(est.usage > 1e8 ? 0 : 1)} MB used` : ''}</small></div></div>
      ${speed.wpm ? `<div class="row"><div class="rt"><b>Your study pace</b><small>${speed.wpm} words a minute${speed.slide ? ` · ${speed.slide}s a slide` : ''}, learnt from your sessions</small></div><button class="btn" data-x="forget">Forget</button></div>` : ''}
      <div class="row"><div class="rt"><b>Remove everything</b><small>Documents, sessions and history</small></div><button class="btn danger" data-x="wipe">Remove</button></div>
    </div>
    <p class="muted" style="font-size:12px;text-align:center;margin:14px 0 4px">Keys: ←/→ pages · Space scrolls then turns · P pause · B break · F flag · G pages · T theme · V layout · +/− zoom</p>`;
  };
  const s = openSheet({ title: 'Settings', body: '' });
  const draw = async () => { const y = s.body.scrollTop; s.body.innerHTML = await render(); s.body.querySelectorAll('input[type=range]').forEach(fillTrack); s.body.scrollTop = y; };
  draw();
  const apply = (k) => {
    savePrefs();
    applyFlowPrefs();
    if (!R.viewer) return;
    if (k === 'view') { reader.classList.toggle('paged-mode', prefs.view !== 'scroll'); R.viewer.setMode(); }
    else if (['fit', 'fontSize', 'flowW', 'notes'].includes(k)) { if (k === 'notes') { R.viewer.lru.clear(); const kk = R.k; R.viewer.k = -1; R.viewer.show(kk, 0); } else R.viewer.relayout(); }
    else if (k === 'hudCompact') renderHud();
  };
  s.body.addEventListener('click', async (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.fx === 'show') { reveal(b.dataset.p); return; }
    if (b.dataset.fx === 'rm') { await removeFolder(b.dataset.p); draw(); return; }
    if (b.dataset.x === 'addFolder') { s.close(true); addFolder(); return; }
    if (b.dataset.x === 'refresh') { syncFolders({ announce: 'refresh' }); return; }
    if (b.dataset.set) { const v = b.dataset.v; prefs[b.dataset.set] = isNaN(Number(v)) ? v : Number(v); apply(b.dataset.set); draw(); }
    else if (b.dataset.x === 'themes') { s.close(true); openThemes(); }
    else if (b.dataset.x === 'forget') { speed = { wpm: null, slide: null }; await store.del('speed'); draw(); }
    else if (b.dataset.x === 'wipe') {
      if (b.dataset.sure) { for (const d of docs) await store.deleteDoc(d.id); await store.del('session'); await store.del('history'); await store.del('speed'); await store.del('folders'); folders = []; docs = []; selected.clear(); s.close(); if (!reader.hidden) exitReader(); renderLibrary(); toast('Everything removed'); }
      else { b.dataset.sure = 1; b.textContent = 'Tap again to remove'; }
    }
  });
  s.body.addEventListener('input', (e) => {
    const t = e.target;
    if (t.dataset.k) { prefs[t.dataset.k] = t.checked; apply(t.dataset.k); if (t.dataset.k === 'wake') keepAwake(!!R.pacer); }
    else if (t.dataset.r) { prefs[t.dataset.r] = Number(t.value); fillTrack(t); t.closest('.row').querySelector('output').textContent = t.dataset.r === 'lineH' ? t.value : `${t.value}px`; apply(t.dataset.r); }
  });
}

// ================= actions & keys =================
document.addEventListener('click', (e) => {
  const b = e.target.closest('[data-act]');
  if (!b || b.closest('.sheet')) return;
  const a = b.dataset.act;
  const acts = {
    import: () => filePick.click(), themes: openThemes, settings: openSettings,
    selectAll: () => { const pool = docsUnder(cwd); const all = pool.every((d) => selected.has(d.id)); pool.forEach((d) => (all ? selected.delete(d.id) : selected.add(d.id))); renderLibrary(); },
    addFolder, refresh: () => syncFolders({ announce: 'refresh' }),
    clearSel: () => { selected.clear(); renderLibrary(); },
    readSel: () => openFree(readingOrder([...selected])),
    paceSel: () => openSetup(readingOrder([...selected])),
    resume: resumeSession,
    endSaved: async () => { const ses = await store.get('session'); if (ses) { const p = Pacer.restore(ses); p.finish(); await learnSpeed(p); } await store.del('session'); renderHero(); toast('Session ended'); },
    search: () => openSearch(), exit: () => exitReader(), prev: prevPage, next: nextPage, pause: togglePause, ring: togglePause,
    flag: toggleFlag, overview: openOverview, stats: () => openStats(), more: () => moreMenu(b), paceThis,
  };
  acts[a]?.();
});
$('#sortSeg').onclick = (e) => { const b = e.target.closest('[data-sort]'); if (!b) return; prefs.sort = b.dataset.sort; savePrefs(); renderLibrary(); };
document.addEventListener('keydown', (e) => {
  if ((e.metaKey || e.ctrlKey) && !e.altKey && e.key.toLowerCase() === 'f' && !sheets.length) { e.preventDefault(); openSearch(); return; }
  if (SV.open) { if (e.key === 'Escape' && !e.target.closest('#q')) { e.preventDefault(); if (menuEl) closeMenu(); else closeSearch(); } return; }
  if (e.target.closest('input, textarea, select, [contenteditable]')) return;
  if (e.key === '/' && !sheets.length && !R.veil) { e.preventDefault(); openSearch(); return; }
  if (e.metaKey || e.ctrlKey || e.altKey) {
    if (!reader.hidden && (e.key === '=' || e.key === '+' || e.key === '-' || e.key === '0')) { e.preventDefault(); showZoom(R.viewer.setZoom(e.key === '0' ? 1 : R.viewer.zoom * (e.key === '-' ? 1 / 1.15 : 1.15))); }
    return;
  }
  if (e.key === 'Escape') {
    if (themeBarOpen()) { openThemes(); return; }
    if (menuEl) { closeMenu(); return; }
    if (sheets.length) { sheets[sheets.length - 1].close(); return; }
    if (R.veil) return;
    if (!reader.hidden) { if (reader.classList.contains('immersive')) reader.classList.remove('immersive'); else exitReader(); }
    return;
  }
  if (sheets.length) return;
  const k = e.key.toLowerCase();
  if (k === 't') { openThemes(); return; }
  if (reader.hidden) { if (k === 'i' || k === 'o') filePick.click(); return; }
  if (R.veil) { if (k === 'enter') R.veil.el.querySelector('.btn.primary')?.click(); return; }
  const acts = {
    arrowright: nextPage, arrowleft: prevPage, pagedown: () => R.viewer.scrollBy(1) || nextPage(), pageup: () => R.viewer.scrollBy(-1) || prevPage(),
    ' ': () => (e.shiftKey ? R.viewer.scrollBy(-1) || prevPage() : R.viewer.scrollBy(1) || nextPage()),
    j: nextPage, k: prevPage, p: togglePause, e: editSession, f: toggleFlag, g: openOverview, s: () => openStats(), v: toggleView,
    b: () => { if (R.pacer) { R.pacer.takeBreakNow(); showBreak(); } },
    home: () => go(0), end: () => go(R.units.length - 1),
    '+': () => showZoom(R.viewer.setZoom(R.viewer.zoom * 1.15)), '=': () => showZoom(R.viewer.setZoom(R.viewer.zoom * 1.15)), '-': () => showZoom(R.viewer.setZoom(R.viewer.zoom / 1.15)),
    h: () => reader.classList.toggle('immersive'),
  };
  if (k === 'arrowdown' || k === 'arrowup') { if (R.viewer.mode === 'scroll') return; if (!R.viewer.scrollBy(k === 'arrowdown' ? 1 : -1)) (k === 'arrowdown' ? nextPage : prevPage)(); e.preventDefault(); return; }
  if (acts[k]) { e.preventDefault(); acts[k](); }
});

// pause when you leave, and say so when you come back
document.addEventListener('visibilitychange', () => {
  const p = R.pacer;
  if (document.visibilityState === 'hidden') {
    if (p && !p.s.finished) { p.tick(); if (prefs.autoPause && !p.s.paused && !p.s.brk) { p.pause(); R.awayPaused = Date.now(); } saveSession(); }
  } else {
    if (p) keepAwake(true);
    if (p && R.awayPaused) {
      const away = (Date.now() - R.awayPaused) / 1000;
      R.awayPaused = false;
      if (away > 4) toast(`Paused while you were away (${human(away)})`, { ms: 8000, actions: [{ label: 'Resume', primary: true, run: () => { if (p.s.paused) togglePause(); } }] });
      else p.resume();
      paintHud(true);
    }
  }
});
addEventListener('pagehide', () => { if (R.pacer) { R.pacer.tick(); saveSession(); } });

// ================= start =================
applyFlowPrefs();
(async () => {
  await loadSpeed();
  await loadFolders();
  await loadLibrary();
  setTimeout(backfillTexts, 2500);
  if (isDesktop) {
    document.body.classList.add('desktop');
    docs.filter((d) => d.pending).forEach((d) => queueIndex(d.id));
    syncFolders();
    let last = Date.now();
    addEventListener('focus', () => { if (Date.now() - last > 15000) { last = Date.now(); syncFolders({ quiet: false }); } });
  }
  // pick up an unfinished session straight away if the app was closed mid-read
  const ses = await store.get('session');
  if (ses && !ses.pacer?.s?.finished && Date.now() - (ses.pacer?.s?.last || 0) < 6 * 3600e3 && new URLSearchParams(location.search).get('resume') !== '0') {
    toast('Pick up where you left off?', { ms: 9000, actions: [{ label: 'Resume', primary: true, run: resumeSession }] });
  }
})();
if ('serviceWorker' in navigator && location.protocol !== 'file:' && !/localhost|127\.0\.0\.1/.test(location.hostname)) navigator.serviceWorker.register('sw.js').catch(() => {});
window.__sp = { R, prefs, store, docs: () => docs };
